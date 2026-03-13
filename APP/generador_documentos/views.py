import io
import os
import pandas as pd
import zipfile
from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.contrib import messages
from django.conf import settings
from docxtpl import DocxTemplate

def formatear_fecha_texto(fecha_raw):
    """ Convierte fecha a formato: 13 de marzo de 2026 """
    if not fecha_raw or pd.isna(fecha_raw):
        return ""
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    try:
        dt = pd.to_datetime(fecha_raw)
        return f"{dt.day} de {meses[dt.month - 1]} de {dt.year}"
    except:
        return str(fecha_raw)

def subir_excel(request):
    """ Vista que recibe el archivo Excel y muestra la lista de personas """
    if request.method == "POST" and "archivo_excel" in request.FILES:
        excel = request.FILES["archivo_excel"]
        if not excel.name.endswith(('.xlsx', '.xls')):
            messages.error(request, "¡Formato inválido! Solo se permiten archivos .xlsx o .xls")
            return redirect('subir_excel_view')

        try:
            df = pd.read_excel(excel)
            # Convertimos a JSON para persistirlo en la sesión
            excel_json = df.to_json(date_format='iso', orient='split')
            
            # Guardamos en la sesión
            request.session['excel_data'] = excel_json
            
            personas = []
            for i, f in df.iterrows():
                nombre = f"{f.get('NOMBRE1','')} {f.get('NOMBRE 2','')} {f.get('APELLIDO1','')} {f.get('APELLIDO 2','')}".replace('nan','').strip().upper()
                personas.append({
                    "index": i, 
                    "nombre": " ".join(nombre.split()), 
                    "identificacion": str(f.get("IDENTIFICACIÓN", f.get("IDENTIFICACION", ""))), 
                    "cargo": str(f.get("CARGO", "")),
                    "direccion": str(f.get("DIRECCION", "")),
                })
            
            return render(request, "seleccionar_persona.html", {
                "personas": personas,
            })
        except Exception as e:
            messages.error(request, f"Error al procesar Excel: {str(e)}")
            return redirect('subir_excel_view')

    return render(request, "subir_excel.html")

def generar_word(request):
    """ Vista que genera el ZIP con los documentos """
    if request.method == "POST":
        try:
            # Recuperamos el JSON de la sesión
            excel_json = request.session.get('excel_data')
            if not excel_json:
                messages.error(request, "No se encontraron datos del Excel en la sesión. Por favor, vuelve a cargar el archivo.")
                return redirect('subir_excel_view')

            # Reconstruimos el DataFrame desde el string JSON
            df = pd.read_json(io.StringIO(excel_json), orient='split')
            idx = int(request.POST.get("persona_index"))
            fila = df.iloc[idx]
            
            nombre_completo = request.POST.get("nombre_completo", "DOCUMENTO").upper()
            tipo_contrato_seleccionado = request.POST.get("tipo_contrato")

            # --- PROCESAMIENTO DE FECHAS ---
            # Convertir nombres de columna a string antes de aplicar .upper()
            col_fecha = next((c for c in df.columns if "INGRESO" in str(c).upper()), None)
            if col_fecha is not None:
                fecha_ingreso_str = formatear_fecha_texto(fila[col_fecha])
            else:
                fecha_ingreso_str = ""
                
            fecha_examenes_str = formatear_fecha_texto(request.POST.get("fecha_examenes"))
            fecha_inicio_labores = formatear_fecha_texto(request.POST.get("fecha_inicio_labores"))
            fecha_terminacion = formatear_fecha_texto(request.POST.get("fecha_terminacion"))

            # --- FORMATEO DE SALARIO ---
            salario_raw = request.POST.get("salario_mensual", "0")
            try:
                salario_limpio = "".join(filter(str.isdigit, str(salario_raw)))
                salario_formateado = "{:,}".format(int(salario_limpio)).replace(",", ".")
            except:
                salario_formateado = salario_raw

            # --- PROCESAMIENTO LUGAR Y FECHA NACIMIENTO ---
            lugar_nac = request.POST.get("lugar_nacimiento", "")
            fecha_nac_raw = request.POST.get("fecha_nacimiento", "")
            fecha_nac_texto = formatear_fecha_texto(fecha_nac_raw)
            nacimiento_detalles = f"{lugar_nac}, {fecha_nac_texto}" if lugar_nac and fecha_nac_texto else (lugar_nac or fecha_nac_texto)

            # --- HARDWARE ---
            especificaciones = request.POST.getlist("especificaciones[]")
            referencias = request.POST.getlist("referencia_hw[]")
            tabla_hardware = []
            for spec, ref in zip(especificaciones, referencias):
                if spec.strip():
                    tabla_hardware.append({'hw': spec.strip(), 'ref': ref.strip()})

            contexto = {
                "nombre_completo": nombre_completo,
                "tipo_documento": request.POST.get("tipo_documento"),
                "identificacion": request.POST.get("identificacion"),
                "cargo": request.POST.get("cargo"),
                "fecha_ingreso": fecha_ingreso_str,
                "ciudad": request.POST.get("ciudad"),
                "horario_trabajo": request.POST.get("horario"),
                "direccion_teletrabajo": request.POST.get("direccion"),
                "fecha_examenes": fecha_examenes_str,
                "recomendaciones": [r for r in request.POST.getlist("recomendaciones[]") if r.strip()],
                "tipo_contrato": tipo_contrato_seleccionado,
                "tabla_hw": tabla_hardware,
                "direccion_empleado": request.POST.get("direccion_empleado"),
                "nacimiento_detalles": nacimiento_detalles,
                "salario_mensual": salario_formateado,
                "eps": request.POST.get("eps"),
                "afp": request.POST.get("afp"),
                "cesantias": request.POST.get("cesantias"),
                "duracion_contrato": request.POST.get("duracion_contrato"),
                "fecha_inicio_labores": fecha_inicio_labores,
                "fecha_terminacion": fecha_terminacion,
                "j_medio": "X" if request.POST.get("jornada") == "Medio Tiempo" else "",
                "j_completo": "X" if request.POST.get("jornada") == "Tiempo Completo" else "",
                "j_otro": "X" if request.POST.get("jornada") == "Otro" else "",
                "j_otro_val": request.POST.get("jornada_otro_texto", ""),
                "u_residencia": "X" if request.POST.get("ubicacion_tipo") == "Residencia" else "",
                "u_otro": "X" if request.POST.get("ubicacion_tipo") == "Otro" else "",
                "u_otro_val": request.POST.get("ubicacion_otro_texto", ""),
            }

            mapa_plantillas = {
                "NDA": "NDA-copia.docx",
                "HOME_OFFICE": "Homme Ofice Agreement-copia.docx",
                "MEDICO": "RECOMENDACIONES MEDICAS-copia.docx",
                "CONTRATO": ""
            }

            if tipo_contrato_seleccionado == "Indefinido":
                mapa_plantillas["CONTRATO"] = "EmploymentContract_Indefinite_Ordinary Salary-copia.docx"
            elif tipo_contrato_seleccionado == "Indefinido Integral":
                mapa_plantillas["CONTRATO"] = "EmploymentContract_Indefinite_Integral Salary-copia.docx"
            elif tipo_contrato_seleccionado == "Fijo Integral":
                mapa_plantillas["CONTRATO"] = "EmploymentContract_FixedTerm_Integral Salary-copia.docx"
            elif tipo_contrato_seleccionado == "Fijo":
                mapa_plantillas["CONTRATO"] = "EmploymentContract_FixedTerm_Ordinary Salary-copia.docx"

            seleccionados = request.POST.getlist("archivos_a_generar")
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for clave in seleccionados:
                    nombre_archivo_plantilla = mapa_plantillas.get(clave)
                    if nombre_archivo_plantilla:
                        ruta_p = os.path.join(settings.BASE_DIR, 'media', 'plantillas', nombre_archivo_plantilla)
                        if os.path.exists(ruta_p):
                            doc = DocxTemplate(ruta_p)
                            doc.render(contexto)
                            output = io.BytesIO()
                            doc.save(output)
                            zip_file.writestr(f"{nombre_completo} - {nombre_archivo_plantilla}", output.getvalue())

            zip_buffer.seek(0)
            if zip_buffer.getbuffer().nbytes < 100:
                 messages.warning(request, "No se pudo generar ningún archivo. Verifica que hayas seleccionado al menos un tipo de documento.")
                 return redirect('subir_excel_view')

            response = HttpResponse(zip_buffer.getvalue(), content_type="application/zip")
            response["Content-Disposition"] = f'attachment; filename="Docs_{nombre_completo.replace(" ","_")}.zip"'
            return response

        except Exception as e:
            messages.error(request, f"Error al generar documentos: {str(e)}")
            return redirect('subir_excel_view')

    # Si alguien entra por GET, redirigir
    return redirect('subir_excel_view')