import os, datetime, json
import customtkinter as ctk
from tksheet import Sheet
from tkinter import messagebox, Menu, simpledialog
from PIL import Image

#Libreria especifica para descarga y carga de informacion
import ONE_decode

# Configuración inicial
ctk.set_appearance_mode("dark")  # Puedes cambiar a "Dark" o "Light"
ctk.set_default_color_theme("blue")

# Crear la ventana principal
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("drONEST Decode")
        self.geometry("900x600")  # Puedes ajustar esto según tu gusto
        self.minsize(600, 400)
        self.grid_rowconfigure(1, weight=2)
        self.grid_columnconfigure(0, weight=1)

        # Crear secciones
        self.frame_changer()
        self.menu_frame()
        #DECODIG FRAME
        self.create_top_frame()
        self.create_bottom_frame()

        self.info_image = None
        self.escribir = {"date": "", "total": 0, "trabajador":"", "nss":"", "array": []}
        self.nombre_documento = ""

        self.pausar_proceso = False
        self.json_autoguardado = {
            "path_fotos": None,
            "fotos_decodificadoas": [],
            "path_saver": None,
            "escribiendo": None
        }

        self.registro_personal()

    def frame_changer(self):
        self.decode_frame = ctk.CTkFrame(self)
        self.decode_frame.grid(row=0, column=0, sticky="nsew")
        self.decode_frame.grid_columnconfigure(0, weight=1)
        self.decode_frame.grid_rowconfigure((0, 1), weight=1)
        self.decode_frame.pack(fill="both", expand=True)

    def menu_frame(self):
        def mostrar_frame_decodificador():
            self.menubar.entryconfig(2, state="disabled")
            self.menubar.entryconfig(3, state="normal")
            try:
                self.decode_frame.pack(fill="both", expand=True)
                self.create_qr_frame.destroy()
            except:pass

        def switch_event():
            if self.switch_var.get() == "0": ctk.set_appearance_mode("dark")
            else: ctk.set_appearance_mode("light")

        def cambiar_usuario():
            mostrar_frame_decodificador()
            self.registro_personal()

        def closer():
            msg = "¿Esta seguro de que desea cerrar la aplicacion?"
            ans = messagebox.askokcancel(title="CERRAR", message=msg)
            if ans == True: self.destroy()
        #MENU
        # 1. Create the main top-level menu bar
        self.menubar = Menu(self)

        # 2. Create a "File" dropdown menu item
        self.file_menu = Menu(self.menubar, tearoff=0)
        self.file_menu.add_command(label="Cambiar de Usuario", command=cambiar_usuario)
        self.switch_var = ctk.StringVar(value="off")
        self.file_menu.add_checkbutton(label="Modo Claro",variable=self.switch_var,command=switch_event)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Cerrar", command=closer)

        # 4. Link the dropdown menus to the main menu bar
        self.menubar.add_cascade(label="Opciones", menu=self.file_menu)
        self.config(menu=self.menubar)

        # 5. Nombre y Nss de trabajador
        try:
            self.my_name = self.menubar.add_command(label=f"            {self.escribir['trabajador']} - {self.escribir['nss']}", state="disable")
            self.my_name.pack(side='Right')
        except:pass

    def registro_personal(self):
        def informacion():
            nombre = name_entry.get()
            nss = nss_entry.get()
            msg = f"¿La informacion del trabajador es correcta?\n\n{nombre}\n{nss}"
            if nombre == '' or nss == '':
                msg = "Nombre/NSS no colocado/s"
                messagebox.showwarning(title="WARNING", message=msg)
            else:
                ans = messagebox.askyesno(title="Registro correcto", message=msg)
                if ans == True:
                    self.escribir['trabajador'] = nombre
                    self.escribir['nss'] = nss
                    self.ventana_escritura.pack_forget()

                    self.menu_frame()

                    self.start_button.configure(state="normal")
                    self.resend_button.configure(state="normal")

        self.start_button.configure(state="disabled")
        self.resend_button.configure(state="disabled")
        try:
            self.ventana_escritura.pack_forget()
        except: pass

        self.ventana_escritura = ctk.CTkFrame(self.bottom_frame)#Toplevel()
        self.ventana_escritura.pack(fill='both', expand=True, padx=5, pady=5)
        #self.ventana_escritura.title("Registro del trabajador")
        #self.ventana_escritura.geometry("400x150")
        #self.ventana_escritura.resizable(False, False)

        frame_sectop=ctk.CTkFrame(self.ventana_escritura)
        frame_sectop.pack(fill='both', expand=True, side='top')
        frame_secbottom=ctk.CTkFrame(self.ventana_escritura)
        frame_secbottom.pack(fill='both', expand=True, side="bottom")

        frame_sec1 = ctk.CTkFrame(frame_sectop)
        frame_sec1.pack(fill='both', expand=True, side='top')
        name_label = ctk.CTkLabel(frame_sec1, text="Nombre de trabajador", font=("Arial", 24, "bold"))
        name_label.pack(fill='both', expand=True, side='left', padx=10, pady=2)
        name_entry = ctk.CTkEntry(frame_sec1)
        name_entry.pack(fill='both', expand=True, side='right', padx=5, pady=2)

        frame_sec2 = ctk.CTkFrame(self.ventana_escritura)
        frame_sec2.pack(fill='both', expand=True, side='bottom')
        nss_label = ctk.CTkLabel(frame_sec2, text="NSS", font=("Arial", 24, "bold"))
        nss_label.pack(fill='both', expand=True, side='left', padx=10, pady=2)
        nss_entry = ctk.CTkEntry(frame_sec2)
        nss_entry.pack(fill='both', expand=True, side='right', padx=5, pady=2)

        boton_registro = ctk.CTkButton(frame_secbottom, text="Registrar", command=informacion, font=("Arial", 24, "bold"))
        boton_registro.pack(fill='both', expand=True, padx=10, pady=2)

        #ONE_decode.excel_process()

    #FRAME DE DECODIFICACION
    def create_top_frame(self):
        # Frame superior - 1/3 de la pantalla
        self.top_frame = ctk.CTkFrame(self.decode_frame)
        self.top_frame.grid(row=0, column=0, sticky="nsew")
        self.top_frame.grid_rowconfigure(0, weight=1)
        self.top_frame.grid_columnconfigure((0, 1, 2), weight=1)  # 3 columnas

        # Cargar imagen (ajusta la ruta a tu imagen)
        try:
            image = ctk.CTkImage(dark_image=Image.open("media\\logo8.png"), size=(450, 200))
            self.image_label = ctk.CTkLabel(self.top_frame, image=image, text="")
            self.image_label.grid(row=0, column=0, padx=10, pady=10)
        except:
            self.image_label = ctk.CTkLabel(self.top_frame, text="Imagen no encontrada")
            self.image_label.grid(row=0, column=0, padx=10, pady=10)

        # Botón central
        self.start_button = ctk.CTkButton(self.top_frame, text="DECODIFICAR", command=self.buscador,
                                          font=("Arial", 24, "bold"))
        self.start_button.configure(height=50, width=200)
        self.start_button.grid(row=0, column=1, padx=10, pady=10)

        # Botón derecho
        self.resend_button = ctk.CTkButton(self.top_frame, text="RE ENVIAR", command=self.re_envio, fg_color="green",
                                           font=("Arial", 24, "bold"))
        self.resend_button.configure(height=50, width=200)
        self.resend_button.grid(row=0, column=2, padx=10, pady=10)

    def create_bottom_frame(self):
        # Frame inferior - 2/3 de la pantalla
        self.bottom_frame = ctk.CTkFrame(self.decode_frame)
        self.bottom_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.bottom_frame.grid_rowconfigure(0, weight=1)
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        info_ruta = "RUTA DE BUSQUEDA\nINFORMACION DECODIFICADA"
        self.info_label = ctk.CTkLabel(self.bottom_frame, text=info_ruta, font=("Arial", 24, "bold") )
        self.info_label.pack(fill="both", expand=True)
        # Scrollable Frame
        self.scroll_frame = ctk.CTkScrollableFrame(self.bottom_frame)


    #########################
    # SECCIONES A DECODIFICAR

    def buscador(self):
        title = "Decodificador drOnest"
        #self.path = ctk.filedialog.askdirectory(title=title)
        a = ctk.filedialog.askopenfilenames(
            title=title,
            filetypes=(
                ("Image Files", "*.JPG"),
                ("Image Files", "*.jpg"),
                ("Image Files", "*.JPEG"),
                ("Image Files", "*.jpeg"),
                ("All Files", "*.*")
            ))

        try:self.limpiar_scroll()
        except:pass

        if a == '':
            msg = "No se seleccionaron fotografias\n"
            messagebox.showwarning(message=msg, title="Warning")
        else:
            ask = str(a[0][::-1])
            n = 0
            for i in ask:
                if i != '/':
                    n += 1
                else:
                    n += 1
                    break
            self.path = a[0][0: len(ask) - n]
            self.info_label.configure(text=f"RUTA A DECODIFICAR\n{self.path}")
            self.mostrador()

    def mostrador(self):
        self.scroll_frame.pack(fill="both", expand=True)
        self.start_button.grid_forget()
        self.resend_button.grid_forget()

        ##########################
        #Mostrar AVISOS
        aviso = "AVISO\n\nRECARGAR\nBaterias - Mandos\nCelulares"
        self.aviso_label = ctk.CTkLabel(self.top_frame, text=aviso, text_color="orange", font=("Arial", 24, "bold"))
        self.aviso_label.grid(row=0, column=1, padx=10, pady=(50, 5), sticky="n")
        self.boton_pausa = ctk.CTkButton(self.top_frame, text="Pausar", font=("Arial", 24, "bold"), command=self.pausado)
        self.boton_pausa.grid(row=0, column=1, padx=10, pady=(0, 50),sticky="s")

        try:
            aviso_image = ctk.CTkImage(dark_image=Image.open("media\\aviso3.png"), size=(225, 225))
            self.aviso_img_label = ctk.CTkLabel(self.top_frame, image=aviso_image, text="")
            self.aviso_img_label.grid(row=0, column=2)
        except:
            self.aviso_img_label = ctk.CTkLabel(self.top_frame, text="Imagen no encontrada")
            self.aviso_img_label.grid(row=0, column=2, padx=10, pady=10)

        self.decodificar()

    def decodificar(self):
        imagenes = os.listdir(self.path)

        num = 0
        for imagen in imagenes:
            self.info_image = ONE_decode.decodificar_imagen(self.path, imagen)

            if self.info_image == None: pass
            else:
                ubicacion = self.info_image['ubicacion']
                tarima = self.info_image['articulo']

                if ubicacion == 0 or tarima == 0:
                    self.IMI(imagen, tarima, ubicacion, self.info_image)
                    if self.info_image != None:
                        self.info_image["number"] = 0
                        self.mostrar_info(imagen)

                if ubicacion != 0 and tarima != 0:
                    num += 1
                    self.info_image["number"] = num
                    self.mostrar_info(imagen)

                self.auto_guardado(imagen)


        title = "decodificacion terminada"
        msg = "Decodificacion terminada con EXITO"
        respuesta = messagebox.showinfo(title=title, message=msg)

        if respuesta == "ok":
            #self.limpiar_scroll()
            self.writer()

    def mostrar_info(self, imagen):
        ################
        #FRAME DE FRAMES
        self.escribir["array"].append(self.info_image)

        self.frame_cuadro = ctk.CTkFrame(self.scroll_frame)
        self.frame_cuadro.pack(side="top", fill="x", expand=True)
        frame_num = ctk.CTkFrame(self.frame_cuadro, height=30, width=15)
        frame_num.pack(side="left", fill="x", expand=True)
        frame_name = ctk.CTkFrame(self.frame_cuadro, height=30)
        frame_name.pack(side="left", fill="x", expand=True)
        frame_tarim = ctk.CTkFrame(self.frame_cuadro, height=30)
        frame_tarim.pack(side="right", fill="x", expand=True)
        frame_ubic = ctk.CTkFrame(self.frame_cuadro, height=30)
        frame_ubic.pack(side="right", fill="x", expand=True)

        label_num = ctk.CTkLabel(frame_num, text=f"{self.info_image["number"]}", font=("Arial", 15, "bold"))
        label_num.place(relx=0.5, rely=0.5, anchor="center")
        label_img = ctk.CTkLabel(frame_name, text=f"{imagen}", font=("Arial", 15, "bold"))
        label_img.place(relx=0.5, rely=0.5, anchor="center")
        label_ubic = ctk.CTkLabel(frame_ubic, text=f"{self.info_image["ubicacion"]}", font=("Arial", 15, "bold"))
        label_ubic.place(relx=0.5, rely=0.5, anchor="center")
        label_tarim = ctk.CTkLabel(frame_tarim, text=f"{self.info_image["articulo"]}", font=("Arial", 15, "bold"))
        label_tarim.place(relx=0.5, rely=0.5, anchor="center")

        # Forzar actualización visual inmediata
        self.scroll_frame.update()

    #########################
    #INTERFACE MALA IMAGEN

    def IMI(self, imagen, tarima, ubicacion, info_image):
        ######################
        #CONTROL DE ENVIOS
        def enviar():
            ubic_chec = info_image['ubicacion']
            tarim_chec = info_image['articulo']

            if tarim_chec == 0:
                tarim_chec = entry_tarima.get()
                if tarim_chec == '': tarim_chec = 0
            if ubic_chec == 0:
                ubic_chec = entry_ubicacion.get()
                if ubic_chec == '': ubic_chec = 0

            ####checador
            text = f"¿Desea guardar la siguiente informacion?\n\nTarima: {tarim_chec}\nUbicacion: {ubic_chec}\n\n"
            complemento = ""
            if str(tarim_chec) == "0": complemento = "La ubicacion esta vacia"
            text = text + complemento
            answer = messagebox.askyesno("Message", text, parent=self.ventana_imi)
            if answer == True:
                info_image['ubicacion'] = ubic_chec
                info_image['articulo'] = tarim_chec

                self.ventana_imi.destroy()
                self.info_image = info_image
                self.decode_frame.pack(fill="both", expand=True)
                #print(info_image)
                #return info_image

        def cerrar():
            text = (
                "Si se cierra la ventana, la informacion de la imagen no se enviara\n¿Desea cerrar de todas formas?")
            answer = messagebox.askyesno("Message", text)#, parent=self.ventana_imi)
            if answer == True:
                self.ventana_imi.destroy()
                self.info_image = None
                self.decode_frame.pack(fill="both", expand=True)

        def noSid():
            ubic_chec = info_image['ubicacion']
            tarim_chec = info_image['articulo']
            if tarim_chec == 0:
                tarim_chec = entry_tarima.get()
                if tarim_chec == '': tarim_chec = 0
            if ubic_chec == 0:
                ubic_chec = entry_ubicacion.get()
                if ubic_chec == '': ubic_chec = 0

            text = ("¿Esta seguro que esta tarima No tiene SID?")
            answer = messagebox.askyesno("Message", text)#, parent=self.ventana_imi)
            if answer == True:
                info_image['ubicacion'] = ubic_chec
                info_image['articulo'] = -1

                self.ventana_imi.destroy()
                self.info_image = info_image
                self.decode_frame.pack(fill="both", expand=True)
                #return info_image

        #Cambiar de frames
        self.decode_frame.pack_forget()
        #EVITAR CONFLICTOS DE FRAMES
        self.menubar.entryconfig(2, state="disabled")
        self.menubar.entryconfig(3, state="disabled")


        ######################
        #Trabajar imagen a decodificar manual
        imagen = Image.open(f"{self.path}\\{imagen}")
        wi, hi = imagen.size
        img_crop = imagen.crop((int(wi / 4), 0, int(wi / 4) * 3, hi))
        img_r = img_crop.resize((int(hi / 5), int(wi / 5)))
        wv, hv = img_r.size

        # Crear ventana secundaria
        self.ventana_imi = ctk.CTkFrame(self, width=wv*3)
        self.ventana_imi.pack(fill="y", expand=True)

        # Cargar imagen
        scroll_frame = ctk.CTkScrollableFrame(self.ventana_imi, width=int(wv*1.5))
        scroll_frame.pack(side='right', fill='both',padx=2)

        my_image = ctk.CTkImage(light_image=img_r, size=(int(wv*1.3), int(hv*1.3)))
        my_label = ctk.CTkLabel(scroll_frame, text="", image=my_image)
        my_label.pack(expand=True)



        # Logo ONEST
        myframe_orden = ctk.CTkFrame(self.ventana_imi, width=240)
        myframe_orden.pack(side='left', fill='both', expand=True, padx=20, pady=2)

        my_logo = ctk.CTkImage(dark_image=Image.open("media\\logo8.png"), size=(240, 160))
        my_label = ctk.CTkLabel(myframe_orden, text="", image=my_logo)
        my_label.place(x=0, y=500)

        fontin = ctk.CTkFont(family='Helvetica', size=16, weight='bold')
        #Crear titulo
        titulo_imi = "Revisar imagenes"
        title_label = ctk.CTkLabel(myframe_orden, text=titulo_imi, font=("Helvetica", 20, "bold"))
        title_label.place(x=35, y=75)

        #crear entradas
        if tarima != 0: text_t = f"Tarima: {tarima}"
        else: text_t = "Escriba aqui QR de tarima"
        entry_info = ctk.CTkLabel(myframe_orden, text=text_t)
        entry_info.place(x=40, y=180)
        entry_tarima = ctk.CTkEntry(master=myframe_orden, placeholder_text=text_t)
        entry_tarima.place(x=40, y=205)

        # Entradas para ubicacion
        if ubicacion != 0:text_u = f"Ubicacion: {ubicacion}"
        else: text_u = "Escriba aqui QR de ubicacion"
        entry_info2 = ctk.CTkLabel(myframe_orden, text=text_u)
        entry_info2.place(x=40, y=255)
        entry_ubicacion = ctk.CTkEntry(master=myframe_orden, placeholder_text=text_u)
        entry_ubicacion.place(x=40, y=280)

        # BOTONES DE CONTROL
        botonSend = ctk.CTkButton(
            master=myframe_orden, text_color="white", command=enviar, height=30, width=100,
            text="SEND", font=fontin, border_color="green", fg_color="green",
            border_width=1)
        botonSend.place(x=65, y=385, anchor='center')

        botonNoSID = ctk.CTkButton(
            master=myframe_orden, text_color="white", command=noSid, height=30, width=100,
            text="No SID", font=fontin, border_color="green", fg_color="gold",
            border_width=1)
        botonNoSID.place(x=120, y=425, anchor='center')

        botonClear = ctk.CTkButton(
            master=myframe_orden, text_color="white", command=cerrar, height=30, width=100,
            text="CLOSE", font=fontin, border_color="black", fg_color="Red",
            border_width=1)
        botonClear.place(x=175, y=385, anchor='center')

        #ESPERAR HASTA QUE TERMINE EL PROCESO
        self.wait_window(self.ventana_imi)

    #########################
    #procesamiento de datos
    def today(self):
        now = datetime.datetime.now()

        y = now.year
        m = now.month
        d = now.day

        hh = now.hour
        mm = now.minute
        ss = now.second

        if d < 10: d = f"0{d}"
        if m < 10: m = f"0{m}"
        if mm < 10: mm = f"0{mm}"
        if ss < 10: ss = f"0{ss}"

        day = [f"{y}-{m}-{d}", f"{hh}-{mm}-{ss}"]
        return day

    def guardado(self):
        now = datetime.datetime.now()
        year = str(now.year)
        month = str(now.month)
        day = str(now.day)

        current = os.getcwd()
        carpetas = os.listdir(current)
        try:
            if 'RC_JSON' not in carpetas:
                os.mkdir(f"{current}\\RC_JSON")
                os.mkdir(f"{current}\\RC_JSON\\{year}")
                os.mkdir(f"{current}\\RC_JSON\\{year}\\{month}")
                os.mkdir(f"{current}\\RC_JSON\\{year}\\{month}\\{day}")
            else:
                carpetas = os.listdir(f"{current}\\RC_JSON")
                if str(year) not in carpetas:
                    os.mkdir(f"{current}\\RC_JSON\\{year}")
                    os.mkdir(f"{current}\\RC_JSON\\{year}\\{month}")
                    os.mkdir(f"{current}\\RC_JSON\\{year}\\{month}\\{day}")
                else:
                    carpetas = os.listdir(f"{current}\\RC_JSON\\{year}")
                    if str(month) not in carpetas:
                        os.mkdir(f"{current}\\RC_JSON\\{year}\\{month}")
                        os.mkdir(f"{current}\\RC_JSON\\{year}\\{month}\\{day}")
                    else:
                        carpetas = os.listdir(f"{current}\\RC_JSON\\{year}\\{month}")
                        if str(day) not in carpetas: os.mkdir(f"{current}\\RC_JSON\\{year}\\{month}\\{day}")

            path_saver = f"{current}\\RC_JSON\\{year}\\{month}\\{day}"
        except:
            path_saver = f"{current}\\RC_JSON\\{year}\\{month}\\{day}"

        return path_saver

    def writer(self):
        path_saver = self.guardado()

        #Nombre para guardar los documentos
        while True:
            self.nombre_documento = simpledialog.askstring("drONEST", "Nombre para el documento")
            if self.nombre_documento == "" or self.nombre_documento == None:
                messagebox.showwarning(title="WARNING", message="No tenemos un nombre para el documento\nFavor de escribir un nombre")
            else:
                break

        #Guardar documento
        day = self.today()
        self.escribir["date"] = f"{day[0]}_{day[1]}"
        self.escribir["total"]= len(self.escribir["array"])
        total = len(self.escribir["array"])
        with open(f"{path_saver}\\{self.nombre_documento} validados{total}  {str(day[1])}.json", 'w') as doc1:
            json.dump(self.escribir, doc1)
            doc1.close()

        #Creacion de reporte excel
        ONE_decode.excel_process(path_saver, self.escribir, self.nombre_documento)

        titulo="drONEST"
        msg = f"Documento guardado con EXITO\n{self.nombre_documento} validados{total}  {str(day[1])}.json\n\nDesea enviar el documento"
        anw = messagebox.askyesno(title=titulo, message=msg)

        if anw == True: ONE_decode.ONE_PROCESS(self.escribir)


        self.normal_states()

    def pausado(self):
        print("Pausar")
        self.pausar_proceso = True
        #self.auto_guardado()

    def auto_guardado(self, name_img):
        if self.pausar_proceso == True:
            msg = "Proceso pausado\n"
            messagebox.showinfo(title="Proceso pausado", message=msg)
        self.pausar_proceso = False

        #print("Auto-Guardado")
        path_fotos = self.path
        foto_actual = name_img
        path_saver = self.guardado()
        auto_guardado = self.escribir
        self.json_autoguardado['path_fotos'] = path_fotos
        self.json_autoguardado['fotos_decodificadoas'].append(foto_actual)
        self.json_autoguardado['path_saver'] = path_saver
        self.json_autoguardado['escribiendo'] = auto_guardado

        #for key in self.json_autoguardado:
        #    print(key, self.json_autoguardado[key])

        fecha = self.today()
        with open(f"{path_saver}\\Autoguardado {fecha[0]}.json", 'w') as doc1:
            json.dump(self.json_autoguardado, doc1)
            doc1.close()

    #########################
    # procesamiento de re-envio de Datos
    def re_envio(self):
        def fotos_almacen(file_path):
            with open(file_path, 'r') as file: data = json.load(file)
            return data

        current = os.getcwd()
        filename = ctk.filedialog.askopenfilename(initialdir=current,
                                              title="Select a File",
                                              filetypes=(("Json Files", "*.json"),
                                                         ("all files", "*.*")))
        if filename == '':
            msg = "Archivo JSON no seleccionado\n"
            messagebox.showwarning(message=msg, title="Warning")
        else:
            escribir = fotos_almacen(filename)
            #print(type(escribir))
            #print(escribir)

            with open(f"{filename}", "w") as j_file:
                json.dump(escribir, j_file, indent=4)
                j_file.close()

            ONE_decode.ONE_PROCESS(escribir)
            self.info_label.configure(text=f"DOCUMENTO RE-ENVIADO\n{filename}")
            for widget in self.scroll_frame.winfo_children():
                widget.destroy()
            self.scroll_frame.pack_forget()

    #########################
    #SECCION PARA LIMPIAR SCROLL
    def normal_states(self):
        try:
            self.menubar.entryconfig(2, state="disabled")
            self.menubar.entryconfig(3, state="normal")
        except: pass

        self.boton_pausa.grid_forget()
        self.aviso_label.grid_forget()
        self.aviso_img_label.grid_forget()
        self.start_button.grid(row=0, column=1, padx=10, pady=10)
        self.resend_button.grid(row=0, column=2, padx=10, pady=10)

    def limpiar_scroll(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self.scroll_frame.pack_forget()
        info_ruta = "RUTA DE BUSQUEDA\nINFORMACION DECODIFICADA"
        self.info_label.configure(text=info_ruta)


if __name__ == "__main__":
    app = App()
    app.mainloop()
