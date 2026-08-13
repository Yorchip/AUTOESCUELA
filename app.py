def save_statistics(estadisticas):
    try:
        datos = []
        for clave, stats in estadisticas.items():
            pregunta_corta = str(stats['pregunta'])[:50].strip()
            datos.append({
                'tema': stats['tema'],
                'pregunta_corta': pregunta_corta,
                'pregunta_completa': stats['pregunta'],
                'intentos': stats['intentos'],
                'aciertos': stats['aciertos'],
                'fallos': stats['fallos'],
                'aciertos_consecutivos': stats.get('aciertos_consecutivos', 0),
                'porcentaje_fallos': round((stats['fallos'] / stats['intentos']) * 100, 1) if stats['intentos'] > 0 else 0,
                'porcentaje_aciertos': round((stats['aciertos'] / stats['intentos']) * 100, 1) if stats['intentos'] > 0 else 0
            })
        
        df = pd.DataFrame(datos)
        if not df.empty:
            df = df.sort_values(['porcentaje_fallos', 'fallos'], ascending=[False, False])
        
        # 1. Guardar localmente en el servidor
        with pd.ExcelWriter(ESTADISTICAS_FILE, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Estadisticas', index=False)
        
        # 2. Guardar en GitHub de forma permanente
        if "GITHUB_TOKEN" in st.secrets:
            from github import Github, GithubException
            token = st.secrets["GITHUB_TOKEN"]
            g = Github(token)
            repo = g.get_repo("Yorchip/AUTOESCUELA")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Estadisticas', index=False)
            
            content_bytes = buffer.getvalue()
            file_path = "estadisticas.xlsx"
            
            # Reintentos automáticos si hay conflicto de versiones (SHA)
            max_intentos = 3
            for intento in range(max_intentos):
                try:
                    contents = repo.get_contents(file_path)
                    repo.update_file(
                        path=file_path,
                        message="📊 Actualización automática de estadísticas",
                        content=content_bytes,
                        sha=contents.sha
                    )
                    break  # Éxito: salimos del bucle
                except GithubException as ge:
                    if ge.status == 404:
                        # Si no existe, lo creamos
                        repo.create_file(
                            path=file_path,
                            message="📊 Creación automática de estadísticas",
                            content=content_bytes
                        )
                        break
                    elif ge.status == 409 and intento < max_intentos - 1:
                        # Si hay conflicto 409, esperamos medio segundo y reintentamos
                        import time
                        time.sleep(0.5)
                        continue
                    else:
                        st.error(f"Error al actualizar estadísticas en GitHub: {ge}")
                        break
    except Exception as e:
        st.error(f"Error al guardar estadísticas: {e}")
