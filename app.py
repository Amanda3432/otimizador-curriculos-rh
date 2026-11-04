        if "|" in linha:
            destaque, complemento = linha.split("|", 1)
            conteudo = f"<b>{texto_para_pdf(destaque.strip())}</b> | {texto_para_pdf(complemento.strip())}"
            return Paragraph(conteudo, estilo_texto)
        return Paragraph(texto_para_pdf(linha), estilo_texto)

    def add_secao_pdf(titulo, primeiro_elemento=None):
        elementos = [
            HRFlowable(width="100%", thickness=0.8, color=cor_total, spaceAfter=2, spaceBefore=3),
            Paragraph(f"<b>{texto_para_pdf(titulo)}</b>", estilo_titulo_secao),
        ]
        if primeiro_elemento is not None:
            elementos.append(primeiro_elemento)
        story.append(KeepTogether(elementos))

    if perfil_txt:
        paragrafos_perfil = dividir_paragrafos_perfil(perfil_txt)[:3]
        if paragrafos_perfil:
            primeiro_perfil = Paragraph(texto_para_pdf(paragrafos_perfil[0]), estilo_perfil_texto)
            add_secao_pdf("Perfil Profissional", primeiro_perfil)
            for p_paragrafo in paragrafos_perfil[1:]:
                story.append(Paragraph(texto_para_pdf(p_paragrafo), estilo_perfil_texto))

    if formacao_txt:
        linhas_formacao_pdf = [linha.strip() for linha in formacao_txt.splitlines() if linha.strip()]
        elementos_formacao = [paragrafo_com_destaque_pdf(linha) for linha in linhas_formacao_pdf]
        if elementos_formacao:
            add_secao_pdf("Formação Acadêmica", elementos_formacao[0])
            story.extend(elementos_formacao[1:])

    if secao_tem_conteudo(cursos_txt):
        linhas_cursos_pdf = [linha.strip() for linha in cursos_txt.splitlines() if linha.strip()]
        elementos_cursos = [paragrafo_com_destaque_pdf(linha) for linha in linhas_cursos_pdf]
        if elementos_cursos:
            add_secao_pdf("Cursos e Certificações", elementos_cursos[0])
            story.extend(elementos_cursos[1:])

    if secao_tem_conteudo(idiomas_txt) and "não constante" not in idiomas_txt.lower():
        linhas_idiomas_pdf = [linha.strip() for linha in idiomas_txt.splitlines() if linha.strip()]
        elementos_idiomas = [Paragraph(texto_para_pdf(linha), estilo_texto) for linha in linhas_idiomas_pdf]
        if elementos_idiomas:
            add_secao_pdf("Idiomas", elementos_idiomas[0])
            story.extend(elementos_idiomas[1:])

    if secao_tem_conteudo(projetos_txt) and "não constante" not in projetos_txt.lower():
        itens_projetos_pdf = [linha.lstrip("-•* ").strip() for linha in projetos_txt.splitlines() if linha.strip()]
        elementos_projetos = [Paragraph(texto_para_pdf(f"• {item}"), estilo_texto) for item in itens_projetos_pdf]
        if elementos_projetos:
            add_secao_pdf("Projetos e Resultados", elementos_projetos[0])
            story.extend(elementos_projetos[1:])

    if secao_tem_conteudo(habilidades_txt):
        itens_habilidades_pdf = [linha.lstrip("-•* ").strip() for linha in habilidades_txt.splitlines() if linha.strip()]
        elementos_habilidades = [Paragraph(texto_para_pdf(f"• {item}"), estilo_texto) for item in itens_habilidades_pdf]
        if elementos_habilidades:
            add_secao_pdf("Habilidades e Competências", elementos_habilidades[0])
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

    if os.path.exists("temp_foto.png"):
        os.remove("temp_foto.png")

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
