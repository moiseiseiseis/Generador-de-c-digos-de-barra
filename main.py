import pandas as pd
import random
import os
import re
from barcode.ean import EAN13
from barcode.writer import ImageWriter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, StringVar

class GeneradorCodigosBarras:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador EAN-13 Avanzado")
        self.root.geometry("750x450")
        
        # Variables
        self.archivo_excel = StringVar()
        self.carpeta_salida = StringVar()
        self.columna_producto = StringVar()
        self.columna_codigo = StringVar()
        
        # Interfaz gráfica
        self.setup_ui()
    
    def setup_ui(self):
        """Configura la interfaz de usuario"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. Selección de archivo Excel
        ttk.Label(main_frame, text="Archivo Excel:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.archivo_excel, width=50).grid(row=0, column=1)
        ttk.Button(main_frame, text="Examinar", command=self.seleccionar_excel).grid(row=0, column=2)
        
        # 2. Selector de columnas
        ttk.Label(main_frame, text="Columna de productos:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.combo_producto = ttk.Combobox(main_frame, textvariable=self.columna_producto, state='readonly')
        self.combo_producto.grid(row=1, column=1, sticky=tk.W)
        
        ttk.Label(main_frame, text="Columna de códigos (opcional):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.combo_codigo = ttk.Combobox(main_frame, textvariable=self.columna_codigo, state='readonly')
        self.combo_codigo.grid(row=2, column=1, sticky=tk.W)
        
        # 3. Carpeta de salida
        ttk.Label(main_frame, text="Carpeta para imágenes:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.carpeta_salida, width=50).grid(row=3, column=1)
        ttk.Button(main_frame, text="Examinar", command=self.seleccionar_carpeta).grid(row=3, column=2)
        
        # 4. Botón de generación
        ttk.Button(main_frame, text="Generar Códigos", 
                 command=self.generar_codigos, style='Accent.TButton').grid(row=4, column=1, pady=20)
        
        # 5. Barra de progreso
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.progress.grid(row=5, column=0, columnspan=3, pady=10)
        
        # Estilo
        style = ttk.Style()
        style.configure('Accent.TButton', foreground='white', background='#28a745')
    
    def seleccionar_excel(self):
        """Selecciona el archivo Excel y carga sus columnas"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            self.archivo_excel.set(archivo)
            try:
                df = pd.read_excel(archivo, nrows=1)  # Lee solo el encabezado
                columnas = df.columns.tolist()
                self.combo_producto['values'] = columnas
                self.combo_codigo['values'] = ['--Ninguna--'] + columnas
                self.combo_producto.current(0)
                self.combo_codigo.current(0)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el archivo:\n{str(e)}")
    
    def seleccionar_carpeta(self):
        """Selecciona la carpeta de salida para las imágenes"""
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta para guardar imágenes")
        if carpeta:
            self.carpeta_salida.set(carpeta)
    
    def calculate_checksum(self, code):
        """Calcula el dígito de control para EAN-13"""
        weights = [3, 1] * 6
        weighted_sum = sum(int(c) * w for c, w in zip(code, weights))
        return (10 - weighted_sum % 10) % 10
    
    def es_ean_valido(self, codigo):
        """Verifica si un código es EAN válido (12/13 dígitos) o comienza con prefijo conocido"""
        if pd.isna(codigo):
            return False
            
        codigo = str(codigo).strip()
        
        # Verificar si es EAN válido (12 o 13 dígitos)
        if codigo.isdigit() and (len(codigo) == 12 or len(codigo) == 13):
            return True
        
        return False
    
    def es_prefijo_valido(self, codigo):
        """Verifica si el código comienza con prefijo de país válido"""
        if pd.isna(codigo):
            return False
            
        codigo = str(codigo).strip()
        
        # Lista de prefijos de país comunes (puedes ampliarla)
        prefijos_paises = [
            '750', # México
            '00', '01', '03', '04', '06', # USA y Canadá
            '30', '31', '32', '33', '34', '35', '36', '37', # Francia
            '40', '41', '42', '43', '44', # Alemania
            '45', '46', '47', # Rusia
            '49', # Japón
            '50', # Reino Unido
            '54', # Bélgica y Luxemburgo
            '57', # Dinamarca
            '64', # Finlandia
            '70', # Noruega
            '76', # Suiza
            '84', # España
            '87', # Holanda
            '93'  # Australia
        ]
        
        return any(codigo.startswith(prefijo) for prefijo in prefijos_paises)
    
    def generar_ean_valido(self):
        """Genera un código EAN-13 válido con prefijo mexicano"""
        prefijo = "750"  # Prefijo para México
        cuerpo = ''.join(random.choices("0123456789", k=9))  # 9 dígitos aleatorios
        codigo_12 = prefijo + cuerpo  # Total: 12 dígitos
        # Añadir dígito de control para convertirlo en EAN-13 válido
        return codigo_12 + str(self.calculate_checksum(codigo_12))
    
    def limpiar_nombre(self, nombre):
        """Limpia el nombre para usarlo en archivos"""
        nombre = str(nombre).strip()
        nombre = re.sub(r'[^\w\s-]', '', nombre)  # Elimina caracteres especiales
        return re.sub(r'\s+', '_', nombre)  # Reemplaza espacios con _
    
    def generar_codigo_barras(self, codigo, nombre_producto, carpeta_salida):
        """Genera la imagen del código de barras"""
        try:
            # Validación final
            if len(codigo) != 13 or not codigo.isdigit():
                raise ValueError(f"Código {codigo} no válido para EAN-13")
            
            # Configurar writer
            writer = ImageWriter()
            writer.set_options({
                'module_width': 0.35,
                'module_height': 15,
                'quiet_zone': 6.5,
                'write_text': False,
                'background': 'white',
                'foreground': 'black'
            })
            
            # Generar código
            ean = EAN13(codigo, writer=writer)
            
            # Nombre de archivo seguro
            nombre_limpio = self.limpiar_nombre(nombre_producto)[:30]
            nombre_archivo = f"ean_{nombre_limpio}_{codigo}.png"
            ruta_completa = os.path.join(carpeta_salida, nombre_archivo)
            
            # Guardar imagen
            ean.save(ruta_completa)
            
            return ruta_completa
            
        except Exception as e:
            print(f"Error generando código {codigo}: {str(e)}")
            return None
    
    def generar_codigos(self):
        """Proceso principal de generación"""
        try:
            # Validación inicial
            if not self.archivo_excel.get() or not self.carpeta_salida.get():
                messagebox.showerror("Error", "Debes seleccionar un archivo Excel y una carpeta de salida")
                return
            
            if not self.columna_producto.get():
                messagebox.showerror("Error", "Debes seleccionar la columna de productos")
                return
            
            # 1. Leer archivo Excel
            df = pd.read_excel(self.archivo_excel.get())
            
            # 2. Procesar códigos existentes
            columna_codigo = self.columna_codigo.get() if self.columna_codigo.get() != "--Ninguna--" else None
            
            if columna_codigo:
               # LÓGICA CORREGIDA - COPIAR DESDE AQUÍ
                df['EAN_FINAL'] = df.apply(
                    lambda row: str(row[columna_codigo]).strip() 
                    if (not pd.isna(row[columna_codigo])) and 
                    (self.es_ean_valido(row[columna_codigo]) or  # Conserva EANs válidos (12/13 dígitos)
                        self.es_prefijo_valido(row[columna_codigo]))  # Conserva prefijos internacionales
                    else self.generar_ean_valido(),  # Genera nuevo para todo lo demás
                    axis=1
                )
                # HASTA AQUÍ
            else:
                # Generar todos nuevos (13 dígitos con checksum)
                df['EAN_FINAL'] = [self.generar_ean_valido() for _ in range(len(df))]
            
            # 3. Validación y limpieza
            df = df.dropna(subset=[self.columna_producto.get()])
            df['EAN_FINAL'] = df['EAN_FINAL'].astype(str).str.strip()
            
            # 4. Preparar carpeta de salida
            carpeta_salida = os.path.abspath(self.carpeta_salida.get())
            os.makedirs(carpeta_salida, exist_ok=True)
            
            # Configurar progreso
            self.progress['maximum'] = len(df)
            self.progress['value'] = 0
            self.root.update_idletasks()
            
            # 5. Generar imágenes solo para códigos EAN nuevos
            resultados = []
            for idx, row in df.iterrows():
                try:
                    codigo = str(row['EAN_FINAL'])
                    nombre = str(row[self.columna_producto.get()])
                    
                    # Solo generar imágenes para códigos nuevos generados (que comienzan con 750)
                    if codigo.startswith('750'):
                        ruta_imagen = self.generar_codigo_barras(codigo, nombre, carpeta_salida)
                        if ruta_imagen:
                            resultados.append(f"Fila {idx+2}: Generado - {os.path.basename(ruta_imagen)}")
                        else:
                            resultados.append(f"Fila {idx+2}: Error al generar imagen")
                    else:
                        resultados.append(f"Fila {idx+2}: Conservado - {codigo}")
                    
                    # Actualizar progreso
                    self.progress['value'] = idx + 1
                    self.root.update_idletasks()
                    
                except Exception as e:
                    resultados.append(f"Fila {idx+2}: Error - {str(e)}")
                    continue
            
            # 6. Guardar Excel actualizado
            nombre_original = os.path.basename(self.archivo_excel.get())
            nuevo_nombre = f"CODIGOS_{nombre_original}"
            ruta_completa = os.path.join(carpeta_salida, nuevo_nombre)
            
            # Ordenar columnas
            columnas = [c for c in df.columns if c != 'EAN_FINAL'] + ['EAN_FINAL']
            df[columnas].to_excel(ruta_completa, index=False)
            
            # 7. Mostrar resumen
            self.mostrar_resumen(resultados, ruta_completa, carpeta_salida)
            
            # Resetear progreso
            self.progress['value'] = 0
            
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error:\n\n{str(e)}")
            self.progress['value'] = 0
    
    def mostrar_resumen(self, resultados, ruta_excel, carpeta_imagenes):
        """Muestra resumen del proceso"""
        exitosos = sum(1 for r in resultados if "Generado" in r)
        conservados = sum(1 for r in resultados if "Conservado" in r)
        errores = sum(1 for r in resultados if "Error" in r)
        
        resumen = (
            f"Proceso completado:\n\n"
            f"• Productos procesados: {len(resultados)}\n"
            f"• Códigos generados: {exitosos}\n"
            f"• Códigos conservados: {conservados}\n"
            f"• Errores: {errores}\n\n"
            f"Archivos generados:\n"
            f"• Excel: {ruta_excel}\n"
            f"• Imágenes: {carpeta_imagenes}\n"
            f"• Log: {os.path.join(carpeta_imagenes, 'log_generacion.txt')}"
        )
        
        # Guardar log
        with open(os.path.join(carpeta_imagenes, 'log_generacion.txt'), 'w') as f:
            f.write("\n".join(resultados))
        
        messagebox.showinfo("Resumen", resumen)


if __name__ == "__main__":
    root = tk.Tk()
    app = GeneradorCodigosBarras(root)
    root.mainloop()