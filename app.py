else:
        dfs = []
        bar = st.progress(0)
        status = st.empty()
        
        for i, f in enumerate(uploaded_files):
            status.text(f"Leyendo {f.name}...")
            df_temp = None # Variable temporal
            
            # Procesar según tipo de archivo
            if f.name.endswith('.xlsx'):
                d = pd.read_excel(f)
                if 'Nota' not in d.columns:
                    d = d.melt(id_vars=[d.columns[0]], var_name="Materia", value_name="Nota")
                    d.columns = ['Alumno', 'Materia', 'Nota']
                df_temp = d
            
            elif f.name.endswith('.pdf'):
                text = extract_text_from_pdf(f)
                # Solo procesamos si hay texto
                if text:
                    df_temp = process_data_with_ai(text, api_key, f.name)
                
            elif 'doc' in f.name:
                text = extract_text_from_docx(f)
                if text:
                    df_temp = process_data_with_ai(text, api_key, f.name)
            
            # --- CORRECCIÓN CRÍTICA AQUÍ ---
            # Solo añadimos a la lista si df_temp NO es None (es decir, si funcionó)
            if df_temp is not None and not df_temp.empty:
                dfs.append(df_temp)
            
            bar.progress((i+1)/len(uploaded_files))
            
        # Comprobamos si la lista dfs tiene algo antes de concatenar
        if dfs:
            full_df = pd.concat(dfs, ignore_index=True)
            res = analyze_data(full_df)
            
            # --- Aquí sigue el resto de tu código de gráficas y Word igual que antes ---
            # Gráficas
            plots = []
            fig, ax = plt.subplots(figsize=(10,4))
            res['stats_materia'].sort_values('Suspensos').plot(
                x='Materia', y=['Aprobados','Suspensos'], kind='bar', stacked=True, 
                color=['#81c784','#e57373'], ax=ax)
            img = io.BytesIO(); plt.savefig(img, format='png'); img.seek(0); plots.append(img)
            
            fig2, ax2 = plt.subplots(figsize=(6,4))
            ax2.bar(['0','1','2','+2'], [res['cero'], res['uno'], res['dos'], res['mas_dos']], color='orange')
            img2 = io.BytesIO(); plt.savefig(img2, format='png'); img2.seek(0); plots.append(img2)
            
            # Word
            doc = generate_word_report(res, plots)
            bio = io.BytesIO(); doc.save(bio); bio.seek(0)
            
            status.success("✅ Informe listo")
            st.download_button("📥 Descargar Word", data=bio, file_name="Informe_1Bach7.docx")
            
            # Resumen visual
            c1, c2, c3 = st.columns(3)
            c1.metric("Promocionan", f"{res['pasan']} ({res['pasan']/(res['total'] or 1):.0%})")
            c2.metric("No promocionan", f"{res['no_pasan']}")
            c3.metric("Nota Media", f"{res['media_global']:.2f}")
            st.dataframe(res['ranking'][['Materia','Suspensos','Pct_Suspensos']])
        else:
            # Mensaje amigable si falló todo
            st.error("❌ No se pudieron extraer datos válidos. Verifica tu API Key o el formato de los archivos.")
