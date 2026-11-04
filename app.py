story.extend(elementos_habilidades[1:])

    if experiencia_txt:
        blocos_exp = [b for b in experiencia_txt.split("\n\n") if b.strip()]
        for indice_bloco, bloco in enumerate(blocos_exp):
            linhas_bloco = [l.strip() for l in bloco.split("\n") if l.strip()]
            if not linhas_bloco:
                continue

            empresa_periodo = linhas_bloco[0]
            if "|" in empresa_periodo:
                partes_emp = empresa_periodo.split("|", 1)
                nome_emp = texto_para_pdf(partes_emp[0].strip())
                resto_emp = texto_para_pdf(partes_emp[1].strip())
                empresa_formatada = f"<b><i>{nome_emp}</i></b> | {resto_emp}"
            else:
                empresa_formatada = f"<b><i>{texto_para_pdf(empresa_periodo)}</i></b>"

            elementos_cargo = [Paragraph(empresa_formatada, estilo_exp_empresa)]
            if len(linhas_bloco) > 1:
                elementos_cargo.append(Paragraph(texto_para_pdf(linhas_bloco[1]), estilo_exp_cargo))
            if len(linhas_bloco) > 2:
                primeiro_item = linhas_bloco[2].lstrip("-•* ").strip()
                elementos_cargo.append(Paragraph(texto_para_pdf(f"• {primeiro_item}"), estilo_texto))

            bloco_mantido = KeepTogether(elementos_cargo)
            if indice_bloco == 0:
                add_secao_pdf("Experiência Profissional", bloco_mantido)
            else:
                story.append(bloco_mantido)

            for item in linhas_bloco[3:]:
                item_limpo = item.lstrip("-•* ").strip()
                story.append(Paragraph(texto_para_pdf(f"• {item_limpo}"), estilo_texto))
            story.append(Spacer(1, 0.5))

    if secao_tem_conteudo(voluntariado_txt) and "não constante" not in voluntariado_txt.lower():
        add_secao_pdf("Trabalho Voluntário")
        story.append(Paragraph(texto_para_pdf(voluntariado_txt).replace('\n', '<br/>'), estilo_texto))

    pdf_doc.build(story)
    buffer_pdf.seek(0)

    # Validação final: nunca disponibiliza PDF acima do limite absoluto de 2 páginas.
    paginas_pdf = len(pypdf.PdfReader(buffer_pdf).pages)
    buffer_pdf.seek(0)

    # Quando LibreOffice estiver instalado no servidor, valida também a paginação real do DOCX.
    paginas_word = None
    conversor_office = shutil.which("libreoffice") or shutil.which("soffice")
    if conversor_office:
        try:
            with tempfile.TemporaryDirectory() as pasta_validacao:
                caminho_docx = os.path.join(pasta_validacao, "curriculo_validacao.docx")
                with open(caminho_docx, "wb") as arquivo_temp:
                    buffer_word.seek(0)
                    arquivo_temp.write(buffer_word.read())
                buffer_word.seek(0)
                subprocess.run(
                    [conversor_office, "--headless", f"-env:UserInstallation=file://{pasta_validacao}/lo_profile",
                     "--convert-to", "pdf", "--outdir", pasta_validacao, caminho_docx],
                    check=True, capture_output=True, timeout=30
                )
                caminho_pdf_word = os.path.join(pasta_validacao, "curriculo_validacao.pdf")
                if os.path.exists(caminho_pdf_word):
                    paginas_word = len(pypdf.PdfReader(caminho_pdf_word).pages)
        except Exception:
            paginas_word = None

    if os.path.exists("temp_foto.png"):
        os.remove("temp_foto.png")

    if paginas_pdf > 2 or (paginas_word is not None and paginas_word > 2):
        st.error(
            f"O currículo ainda ultrapassou o limite de 2 páginas "
            f"(PDF: {paginas_pdf} página(s); Word: {paginas_word if paginas_word is not None else 'não validado'}). "
            "Os arquivos não serão disponibilizados com mais de 2 páginas. A ferramenta mantém o limite rígido para não entregar um currículo fora do padrão. "
            "Reduza informações secundárias e gere novamente."
        )
        st.stop()
    elif paginas_word is None:
        st.info(
            f"Paginação do PDF validada: {paginas_pdf} página(s). "
            "O servidor não possui LibreOffice para conferir automaticamente a paginação do Word; "
            "o DOCX recebeu os mesmos cortes de conteúdo e as mesmas configurações A4/fonte 11."
        )
    else:
        st.success(f"Paginação validada: PDF com {paginas_pdf} página(s) e Word com {paginas_word} página(s).")

    # --- BOTÕES DE DOWNLOAD ---
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            label=f"📥 Baixar em Word ({modelo_escolhido})",
            data=buffer_word,
            file_name="curriculo_otimizado.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    with col_d2:
        st.download_button(
            label=f"📥 Baixar em PDF ({modelo_escolhido})",
            data=buffer_pdf,
            file_name="curriculo_otimizado.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )
