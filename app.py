if os.path.exists("temp_foto.png"):
        os.remove("temp_foto.png")

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
