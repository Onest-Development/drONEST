import qrcode, code128, cv2, os, datetime, requests
import aspose.words as aw

from PIL import Image
from pyzbar import pyzbar
from tkinter import messagebox

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.drawing.image import Image

###########################
#Proceso Envio de INFO A ONE
def ONE_PROCESS(escribir):
    json_escribir = json_cesar(escribir)
    if json_escribir != None: enviar_datos(json_escribir)
def almacen():
    breaker = False
    while breaker == False:
        try:
            url_servicio = "http://10.249.249.3/onews/ords/onewms/almacen/consulta_layout"
            parametros = {"cuenta_id": 335}

            respuesta = requests.post(url_servicio, json=parametros)
            datos = respuesta.json()

            breaker = True
            return datos
        except:
            text = "La computadora no se encuentra conectadas al servidor\n\n¿Desea volver a intentar enviar el documento?"
            ans = messagebox.askokcancel("Error de envio", text)
            if ans == False:
                breaker = True
                return None
def json_cesar(escribir):
    datos = almacen()
    if datos == None:
        return datos
    else:
        identificador = datos["IdentificadorAlmacen"]
        nombre_almacen = datos["NombreAlmacen"]
        almacen_info = datos["Dato"]

        datos_fotos = escribir

        resultado_cesar = {
            "IdentificadorAlmacen": str(identificador),
            "NombreAlmacen": str(nombre_almacen),
            "fecha_transmision": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "resultado": []
        }

        for info_fotos in datos_fotos["array"]:
            tarim_foto = info_fotos["articulo"]
            ubic_foto = info_fotos["ubicacion"]

            for info_almacen in almacen_info:

                ubicacion_almacen = info_almacen['ubicacion_id']
                pasillo_id = info_almacen['pasillo_id']
                lado_id = info_almacen['lado_id']
                posx = info_almacen['pos_x']
                posy = info_almacen['pos_y']

                resultado_item = {
                    "ubicacion_id": str(ubicacion_almacen),
                    "pasillo_id": str(pasillo_id),
                    "lado_id": str(lado_id),
                    "posx": str(posx),
                    "posy": str(posy),
                    "identificador_teorico": None,
                    "identificador_drone": None,
                    "ubicacion_drone": None,
                    "status": None,
                    "error": None
                }

                if str(ubicacion_almacen) == str(ubic_foto):
                    tarimas_almacen = info_almacen["datos_inventario"]
                    for tarima_almacen in tarimas_almacen:
                        tarim_almacen = tarima_almacen["identificador"]
                        if str(tarim_foto) == str(tarim_almacen):

                            if resultado_item not in resultado_cesar["resultado"]:
                                resultado_cesar["resultado"].append(resultado_item)
                                resultado_item["identificador_teorico"] = str(tarim_almacen)
                                resultado_item["identificador_drone"] = str(tarim_foto)
                                resultado_item["ubicacion_drone"] = str(ubic_foto)
                                resultado_item["status"] = "true"
                                resultado_item["error"] = "exito"

                                if resultado_item not in resultado_cesar["resultado"]:
                                    resultado_cesar["resultado"].append(resultado_item)
                            else:
                                resultado_item["identificador_teorico"] = str(tarim_almacen)
                                resultado_item["identificador_drone"] = str(tarim_foto)
                                resultado_item["ubicacion_drone"] = str(ubic_foto)
                                resultado_item["status"] = "true"
                                resultado_item["error"] = "exito"
                                if resultado_item not in resultado_cesar["resultado"]:
                                    resultado_cesar["resultado"].append(resultado_item)
                        else:
                            # TARIMA VACIA VERDADERA
                            # Falso Positivos
                            if str(tarim_foto) == "0":
                                resultado_item["identificador_teorico"] = str(tarim_almacen)
                                resultado_item["identificador_drone"] = str(tarim_foto)
                                resultado_item["ubicacion_drone"] = str(ubic_foto)
                                resultado_item["status"] = "true"
                                resultado_item["error"] = "exito"

                                if resultado_item not in resultado_cesar["resultado"]:
                                    resultado_cesar["resultado"].append(resultado_item)

                            # TARIMA NO SID
                            # Falso Negativo
                            elif str(tarim_foto) == "-1":
                                resultado_item["identificador_teorico"] = str(tarim_almacen)
                                resultado_item["identificador_drone"] = "0"
                                resultado_item["ubicacion_drone"] = str(ubic_foto)
                                resultado_item["status"] = "false"
                                resultado_item["error"] = "no encontrado"

                                if resultado_item not in resultado_cesar["resultado"]:
                                    resultado_cesar["resultado"].append(resultado_item)

                            # TARIMA drone Diferente a la de ONE
                            else:
                                resultado_item["identificador_teorico"] = str(tarim_almacen)
                                resultado_item["identificador_drone"] = str(tarim_foto)
                                resultado_item["ubicacion_drone"] = str(ubic_foto)
                                resultado_item["status"] = "false"
                                resultado_item["error"] = "no encontrado"

                                if resultado_item not in resultado_cesar["resultado"]:
                                    resultado_cesar["resultado"].append(resultado_item)

        return resultado_cesar
def enviar_datos(escribir):#json_cesar_enviar):

    # URL del servicio HTTP
    #url_servicio = "http://10.249.249.3/onews/ords/onewms/almacen/drone_inventario"
    #url_servicio = "http://10.249.249.2:8777/ords/onewms/almacen/drone_inventario_directo"
    #url_servicio = "http://10.249.249.3/onews/ords/onewms/almacen/drone_inventario_directo"

    """web server agregado por WS con nuevas modificaciones     11 Sep 2025"""
    #url_servicio = "http://10.249.249.2:8777/ords/onewms/almacen/drone_confirma_inventario"
    url_servicio = "http://10.249.249.3/ordsprd/onewms/almacen/drone_confirma_inventario"
    #resultado_json = JSON.dumps(json_cesar_enviar)

    #resultado_json = JSON.dumps(escribir)
    resultado_json = escribir

    #my_format = open("RC_FORMAT.json", "r")

    try:
        respuesta = requests.post(url_servicio, json=resultado_json)
        print("RESPUESTA")
        print(respuesta)
        respuesta_json = respuesta.json()
        #print("RESPUESTA")
        #print(respuesta_json)
        msg_correcto = respuesta_json["Mensaje_error"]
        messagebox.showinfo("Envio Correcto", msg_correcto)
        #print("#######################")
        #print(msg_correcto)
    except:
        msg_correcto = f"Error en envio"
        messagebox.showinfo("Error en envio", msg_correcto)

def excel_process(path, escribir, nombre_doc):

    wb = openpyxl.Workbook()
    ws = wb.active

    """FORMATO TITULAR"""
    # Combinar rango de celdas
    ws.merge_cells('A1:I4')

    logo = 'media\\logo8.png'
    img = Image(logo)
    img.width = 250   # Width in pixels
    img.height = 70  # Height in pixels

    ws.add_image(img, 'D1')

    # NOMBRES Y TRABAJADORES
    ws.merge_cells('A5:B5')
    ws.merge_cells('A6:B6')

    ws.merge_cells('C5:I5')
    ws.merge_cells('C6:I6')

    celda_nombre = ws['A5']
    celda_nombre.value = 'Trabajador'
    celda_nombre.alignment = Alignment(horizontal='center', vertical='center')
    celda_nombre.font = Font(size=12, bold=True)
    celda_nombre = ws['C5']
    celda_nombre.value = escribir['trabajador']
    celda_nombre.alignment = Alignment(horizontal='center', vertical='center')


    celda_nss = ws['A6']
    celda_nss.value = 'NSS'
    celda_nss.alignment = Alignment(horizontal='center', vertical='center')
    celda_nss.font = Font(size=12, bold=True)
    celda_nombre = ws['C6']
    celda_nombre.value = escribir['nss']
    celda_nombre.alignment = Alignment(horizontal='center', vertical='center')


    # TOTAL DE IMAGENES / PASILLO
    ws.merge_cells('A8:B8')

    ws.merge_cells('D8:E8')
    ws.merge_cells('G8:H8')

    celda_total = ws['A8']
    celda_total.value = 'Total de Imagenes'
    celda_total.alignment = Alignment(horizontal='center', vertical='center')
    celda_total.font = Font(size=12, bold=True)
    celda_total = ws['C8']
    celda_total.value = len(escribir['array'])
    celda_total.alignment = Alignment(horizontal='center', vertical='center')


    celda_tarim_empty = ws['D8']
    celda_tarim_empty.value = 'Ubicaciones Vacias'
    celda_tarim_empty.alignment = Alignment(horizontal='center', vertical='center')
    celda_tarim_empty.font = Font(size=12, bold=True)

    celda_tarim_null = ws['G8']
    celda_tarim_null.value = "Tarimas sin Etiqueta"
    celda_tarim_null.alignment = Alignment(horizontal='center', vertical='center')
    celda_tarim_null.font = Font(size=12, bold=True)

    "TITULOS"
    ws.merge_cells('A10:C10')
    ws.merge_cells('D10:F10')
    ws.merge_cells('G10:I10')

    celda_nombre_imagen = ws['A10']
    celda_nombre_imagen.value = 'Nombre Imagen'
    celda_nombre_imagen.alignment = Alignment(horizontal='center', vertical='center')
    celda_nombre_imagen.font = Font(size=12, bold=True)

    celda_nombre_tarima = ws['D10']
    celda_nombre_tarima.value = 'Tarima'
    celda_nombre_tarima.alignment = Alignment(horizontal='center', vertical='center')
    celda_nombre_tarima.font = Font(size=12, bold=True)

    celda_nombre_ubicacion = ws['G10']
    celda_nombre_ubicacion.value = 'Ubicacion'
    celda_nombre_ubicacion.alignment = Alignment(horizontal='center', vertical='center')
    celda_nombre_ubicacion.font = Font(size=12, bold=True)

    ubicaciones_null = 0
    tarimas_empty = 0
    for i in range(0, len(escribir['array'])):
        obj = escribir['array'][i]
        num = 11 + i
        ws.merge_cells(f'A{num}:C{num}')
        ws.merge_cells(f'D{num}:F{num}')
        ws.merge_cells(f'G{num}:I{num}')

        celda_nombre_imagen = ws[f'A{num}']
        celda_nombre_imagen.value = obj['img_name']
        celda_nombre_imagen.alignment = Alignment(horizontal='center', vertical='center')

        celda_nombre_imagen = ws[f'D{num}']

        if obj['articulo'] == 0:
            ubicaciones_null += 1
            celda_nombre_imagen.value = "VACIO"
        elif obj['articulo'] == -1:
            tarimas_empty += 1
            celda_nombre_imagen.value = "NO SID"
        else:
            celda_nombre_imagen.value = obj['articulo']

        celda_nombre_imagen.alignment = Alignment(horizontal='center', vertical='center')

        celda_nombre_imagen = ws[f'G{num}']
        celda_nombre_imagen.value = obj['ubicacion']
        celda_nombre_imagen.alignment = Alignment(horizontal='center', vertical='center')

    celda_tarim_empty = ws['F8']
    celda_tarim_empty.value = ubicaciones_null
    celda_tarim_empty.alignment = Alignment(horizontal='center', vertical='center')

    celda_tarim_null = ws['I8']
    celda_tarim_null.value = tarimas_empty
    celda_tarim_null.alignment = Alignment(horizontal='center', vertical='center')

    wb.abs = wb.save(f'{path}\\{nombre_doc}.xlsx')

    buscando = os.listdir(path)
    for doc in buscando:
        if 'Autoguardado' in doc:
            os.remove(f'{path}\\{doc}')
            break

###########################
#Proceso de decodificacion de IMAGENES
def decodificar_imagen(path, imagen):
    formatos = ['.jpg', '.JPG', '.png', '.PNG']
    es_imagne = imagen[-4:]
    if str(es_imagne) not in formatos:
        print("Esta no paso")
        return None
    else:
        #print("->", imagen)
        img = cv2.imread(f"{path}\\{imagen}")
        #print(img.shape)
        h, w, c = img.shape
        img_crop = img[0:h, int(w / 4):int(w / 4) * 3]
        final_img = img_crop

        detect = cv2.QRCodeDetector()
        info = detect.detectAndDecodeMulti(final_img)

        QRS = info[1]
        chodsQRS = info[2]

        info_imagen = {
            "img_name": imagen,
            "ubicacion": 0,
            "articulo": 0,
            "validation": False,
            "datecheck": False,
            "number": 0
        }

        if QRS != None:
            if len(QRS) == 2:
                for pos in range(0, 2):
                    QR = QRS[pos]
                    if QR == '':
                        xs, ys = [], []
                        for chords in chodsQRS[pos]:
                            xs.append(int(chords[0]))
                            ys.append(int(chords[1]))
                        xf, xi = max(xs), min(xs)
                        yf, yi = max(ys), min(ys)
                        try:
                            qr_img = final_img[yi - 50:yf + 50, xi - 50:xf + 50]
                            img_g = cv2.cvtColor(qr_img, cv2.COLOR_BGR2GRAY)  # -imagen con filtro
                            img_ab = cv2.convertScaleAbs(img_g, alpha=1, beta=20)
                            im_r = cv2.resize(img_ab, (600, 600))
                            codes = pyzbar.decode(im_r)
                            if codes:
                                code = codes[0]
                                QR = str(code[0])[2:-1]
                                if len(QR) > 6: info_imagen["articulo"] = QR
                                if len(QR) <= 5: info_imagen["ubicacion"] = QR
                        except:
                            pass
                    else:
                        if len(QR) > 6: info_imagen["articulo"] = QR
                        if len(QR) <= 5: info_imagen["ubicacion"] = QR
                # REVISAR QUEIMAGENES CON QRTARIMA, TENGAN UBICACION
                if info_imagen["ubicacion"] == 0 and info_imagen["articulo"] != 0:
                    ######################
                    # Comienza el filtro para cortar
                    for y in range(4, 0, -1):
                        hc = int(h)
                        yi = 0 + int((h / 6) * (y))
                        yf = int(yi + hc)

                        img_cf = final_img[yi:yf, 0:w]

                        detect = cv2.QRCodeDetector()
                        info_crop = detect.detectAndDecodeMulti(img_cf)

                        QRS = info_crop[1]
                        if QRS:
                            if QRS[0] != '' and len(QRS[0]) <= 6:
                                info_imagen['ubicacion'] = QRS[0]
                                break
                        else:
                            codes = pyzbar.decode(img_cf)
                            if codes:
                                code = codes[0]
                                QR_code = str(code[0])[2:-1]
                                if QR_code != '' and len(QR_code) <= 6:
                                    info_imagen['ubicacion'] = QR_code
                                    break

            if len(QRS) <= 1:
                codes = pyzbar.decode(final_img)
                for code in codes:
                    QR = str(code[0])[2:-1]
                    if len(QR) > 6: info_imagen["articulo"] = QR
                    if len(QR) <= 5: info_imagen["ubicacion"] = QR
                # REVISAR QUEIMAGENES CON QRTARIMA, TENGAN UBICACION
                if info_imagen["ubicacion"] == 0 and info_imagen["articulo"] != 0:
                    ######################
                    # Comienza el filtro para cortar
                    for y in range(4, 0, -1):
                        hc = int(h)
                        yi = 0 + int((h / 6) * (y))
                        yf = int(yi + hc)

                        img_cf = final_img[yi:yf, 0:w]

                        detect = cv2.QRCodeDetector()
                        info_crop = detect.detectAndDecodeMulti(img_cf)

                        QRS = info_crop[1]
                        if QRS:
                            if QRS[0] != '' and len(QRS[0]) <= 6:
                                info_imagen['ubicacion'] = QRS[0]
                                break
                        else:
                            codes = pyzbar.decode(img_cf)
                            if codes:
                                code = codes[0]
                                QR_code = str(code[0])[2:-1]
                                if QR_code != '' and len(QR_code) <= 6:
                                    info_imagen['ubicacion'] = QR_code
                                    break
                # REVISAR CUANDO HAY UBICACION PERO NO ENCONTRO TARIMA
                if info_imagen["ubicacion"] != 0 and info_imagen["articulo"] == 0:
                    for y in range(0, 4):
                        hc = int(h)
                        yi = 0 + int((h / 6) * (y))
                        yf = int(hc - (hc / 2) + yi)

                        img_cf = final_img[yi:yf, 0:w]

                        detect = cv2.QRCodeDetector()
                        info_crop = detect.detectAndDecode(img_cf)

                        QR = info_crop[0]
                        chodsQRS = info_crop[1]
                        if QR:
                            if QR != '' and len(QR) <= 6: info_imagen['ubicacion'] = QR
                            if QR != '' and len(QR) > 7: info_imagen['articulo'] = QR

                            if QR == '':
                                xs, ys = [], []
                                for chords in chodsQRS[0]:
                                    xs.append(int(chords[0]))
                                    ys.append(int(chords[1]))
                                xf, xi = max(xs), min(xs)
                                yf, yi = max(ys), min(ys)

                                qr_img = img_cf[yi - 50:yf + 50, xi - 50:xf + 50]
                                img_g = cv2.cvtColor(qr_img, cv2.COLOR_BGR2GRAY)  # -imagen con filtro
                                img_ab = cv2.convertScaleAbs(img_g, alpha=1, beta=20)
                                im_r = cv2.resize(img_ab, (600, 600))

                                codes = pyzbar.decode(im_r)
                                if codes:
                                    code = codes[0]
                                    QR = str(code[0])[2:-1]
                                    if len(QR) > 6: info_imagen['articulo'] = QR
                                    if len(QR) <= 5: info_imagen['ubicacion'] = QR
                        else:
                            codes = pyzbar.decode(img_cf)
                            if codes:
                                code = codes[0]
                                QR_code = str(code[0])[2:-1]
                                if QR_code != '' and len(QR_code) <= 6:
                                    info_imagen['ubicacion'] = QR_code
                # REVISAR CUANDO NO HAY TARIMA NI UBICACION
                if info_imagen["ubicacion"] == 0 and info_imagen["articulo"] == 0:
                    for y in range(0, 4):
                        hc = int(h)
                        yi = 0 + int((h / 6) * (y))
                        yf = int(hc - (hc / 2) + yi)

                        img_cf = final_img[yi:yf, 0:w]

                        detect = cv2.QRCodeDetector()
                        info_crop = detect.detectAndDecode(img_cf)

                        QR = info_crop[0]
                        chodsQRS = info_crop[1]
                        if QR:
                            if QR != '' and len(QR) <= 6: info_imagen['ubicacion'] = QR
                            if QR != '' and len(QR) > 7: info_imagen['articulo'] = QR

                            if QR == '':
                                xs, ys = [], []
                                for chords in chodsQRS[0]:
                                    xs.append(int(chords[0]))
                                    ys.append(int(chords[1]))
                                xf, xi = max(xs), min(xs)
                                yf, yi = max(ys), min(ys)

                                qr_img = img_cf[yi - 50:yf + 50, xi - 50:xf + 50]
                                img_g = cv2.cvtColor(qr_img, cv2.COLOR_BGR2GRAY)  # -imagen con filtro
                                img_ab = cv2.convertScaleAbs(img_g, alpha=1, beta=20)
                                im_r = cv2.resize(img_ab, (600, 600))

                                codes = pyzbar.decode(im_r)
                                if codes:
                                    code = codes[0]
                                    QR = str(code[0])[2:-1]
                                    if len(QR) > 6: info_imagen['articulo'] = QR
                                    if len(QR) <= 5: info_imagen['ubicacion'] = QR
                        else:
                            codes = pyzbar.decode(img_cf)
                            if codes:
                                code = codes[0]
                                QR_code = str(code[0])[2:-1]
                                if QR_code != '' and len(QR_code) <= 6:
                                    info_imagen['ubicacion'] = QR_code
                    for y in range(3, -1, -1):
                        hc = int(h)
                        yi = 0 + int((h / 6) * (y))
                        yf = int(hc - (hc / 2) + yi)

                        img_cf = final_img[yi:yf, 0:w]

                        detect = cv2.QRCodeDetector()
                        info_crop = detect.detectAndDecode(img_cf)

                        QR = info_crop[0]
                        chodsQRS = info_crop[1]
                        if QR:
                            if QR != '' and len(QR) <= 6: info_imagen['ubicacion'] = QR
                            if QR != '' and len(QR) > 7: info_imagen['articulo'] = QR

                            if QR == '':
                                xs, ys = [], []
                                for chords in chodsQRS[0]:
                                    xs.append(int(chords[0]))
                                    ys.append(int(chords[1]))
                                xf, xi = max(xs), min(xs)
                                yf, yi = max(ys), min(ys)

                                qr_img = img_cf[yi - 50:yf + 50, xi - 50:xf + 50]
                                img_g = cv2.cvtColor(qr_img, cv2.COLOR_BGR2GRAY)  # -imagen con filtro
                                img_ab = cv2.convertScaleAbs(img_g, alpha=1, beta=20)
                                im_r = cv2.resize(img_ab, (600, 600))

                                codes = pyzbar.decode(im_r)
                                if codes:
                                    code = codes[0]
                                    QR = str(code[0])[2:-1]
                                    if len(QR) > 6: info_imagen['articulo'] = QR
                                    if len(QR) <= 5: info_imagen['ubicacion'] = QR
                        else:
                            codes = pyzbar.decode(img_cf)
                            if codes:
                                code = codes[0]
                                QR_code = str(code[0])[2:-1]
                                if QR_code != '' and len(QR_code) <= 6:
                                    info_imagen['ubicacion'] = QR_code

        cv2.destroyAllWindows()
        return info_imagen

    # print("################################################")

