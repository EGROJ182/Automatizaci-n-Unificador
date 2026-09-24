import os
import sys
import logging
from pathlib import Path
from typing import List, Dict, Optional, Union
from enum import Enum
from abc import ABC, abstractmethod
from dataclasses import dataclass
import threading
import pandas as pd
from tkinter import Tk, filedialog, messagebox, ttk, Frame, Label, Button, StringVar, Text
from tkinterdnd2 import DND_FILES, TkinterDnD


# ============================================================================
# CONFIGURACIÓN DE LOGGING
# ============================================================================
class LoggerConfigurator:
    """Configurador centralizado de logging."""
    
    @staticmethod
    def setup(log_file: Optional[str] = None) -> logging.Logger:
        """Configura logging con formato estandarizado."""
        logger = logging.getLogger('ExcelUnifier')
        
        # Evitar duplicados
        if logger.hasHandlers():
            return logger
        
        logger.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        if log_file:
            file_handler = logging.FileHandler(log_file)
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        return logger


# ============================================================================
# EXCEPCIONES PERSONALIZADAS
# ============================================================================
class ExcelUnifierException(Exception):
    """Excepción base para ExcelUnifier."""
    pass


class InvalidFileException(ExcelUnifierException):
    """Levantada cuando un archivo no es válido."""
    pass


class InvalidFolderException(ExcelUnifierException):
    """Levantada cuando una carpeta no existe o no es válida."""
    pass


class UnificationException(ExcelUnifierException):
    """Levantada cuando falla el proceso de unificación."""
    pass


class EncodingException(ExcelUnifierException):
    """Levantada cuando hay problemas de codificación."""
    pass


# ============================================================================
# ENUMERACIONES
# ============================================================================
class FileType(Enum):
    """Tipos de archivo soportados."""
    XLSX = '.xlsx'
    XLS = '.xls'
    CSV = '.csv'


class UnificationMode(Enum):
    """Modos de unificación disponibles."""
    ALL_SHEETS_ALL_FILES = "Todas las hojas de todos los archivos"
    ALL_FILES_MERGE = "Consolidar todos los archivos en una sola hoja"
    BY_FILE = "Por archivo (mantener estructura)"


# ============================================================================
# CLASES DE DATOS
# ============================================================================
@dataclass
class UnificationProgress:
    """Información de progreso durante la unificación."""
    current_file: int
    total_files: int
    percentage: float
    status: str


# ============================================================================
# INTERFAZ Y ABSTRACCIONES
# ============================================================================
class FileReader(ABC):
    """Interfaz para lectores de archivos."""
    
    @abstractmethod
    def read(self, filepath: str, encoding: Optional[str] = None) -> List[Dict]:
        """Lee el archivo y retorna lista de DataFrames con metadatos."""
        pass
    
    @abstractmethod
    def supports(self, file_type: FileType) -> bool:
        """Verifica si soporta el tipo de archivo."""
        pass


# ============================================================================
# IMPLEMENTACIONES DE LECTORES
# ============================================================================
class ExcelFileReader(FileReader):
    """Lector para archivos Excel (.xlsx, .xls)."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
    
    def supports(self, file_type: FileType) -> bool:
        return file_type in [FileType.XLSX, FileType.XLS]
    
    def read(self, filepath: str, encoding: Optional[str] = None) -> List[Dict]:
        """Lee todas las hojas de un archivo Excel."""
        try:
            self.logger.info(f"Leyendo Excel: {filepath}")
            excel_file = pd.ExcelFile(filepath)
            dfs = []
            
            for sheet_name in excel_file.sheet_names:
                df = excel_file.parse(sheet_name, dtype=str)
                dfs.append({
                    'data': df,
                    'sheet_name': sheet_name,
                    'row_count': len(df)
                })
            
            self.logger.debug(f"Excel {filepath} tiene {len(dfs)} hojas")
            return dfs
            
        except Exception as e:
            self.logger.error(f"Error leyendo Excel {filepath}: {str(e)}")
            raise InvalidFileException(f"No se pudo leer {filepath}: {str(e)}")


class CSVFileReader(FileReader):
    """Lector para archivos CSV con soporte multiencodificación."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
    
    def supports(self, file_type: FileType) -> bool:
        return file_type == FileType.CSV
    
    def read(self, filepath: str, encoding: Optional[str] = None) -> List[Dict]:
        """Lee CSV con manejo automático de codificación."""
        encodings_to_try = [encoding] if encoding else self.encodings
        
        for enc in encodings_to_try:
            try:
                self.logger.info(f"Leyendo CSV {filepath} con encoding {enc}")
                df = pd.read_csv(filepath, encoding=enc, on_bad_lines='skip', dtype=str)
                
                self.logger.debug(f"CSV {filepath} leído exitosamente con {enc}")
                return [{
                    'data': df,
                    'sheet_name': Path(filepath).stem,
                    'row_count': len(df)
                }]
                
            except UnicodeDecodeError:
                self.logger.debug(f"Encoding {enc} no funcionó, intentando siguiente...")
                continue
        
        raise EncodingException(
            f"No se pudo leer {filepath} con ninguno de los encodings: {encodings_to_try}"
        )


# ============================================================================
# VALIDADOR DE ARCHIVOS
# ============================================================================
class FileValidator:
    """Valida archivos para procesamiento."""
    
    SUPPORTED_EXTENSIONS = {ft.value for ft in FileType}
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
    
    def validate_file(self, filepath: str) -> bool:
        """Valida que el archivo sea soportado y accesible."""
        try:
            path = Path(filepath)
            
            if not path.exists():
                self.logger.warning(f"Archivo no existe: {filepath}")
                return False
            
            if not path.is_file():
                self.logger.warning(f"No es un archivo: {filepath}")
                return False
            
            if path.suffix.lower() not in self.SUPPORTED_EXTENSIONS:
                self.logger.warning(f"Extensión no soportada: {path.suffix}")
                return False
            
            if path.stat().st_size == 0:
                self.logger.warning(f"Archivo vacío: {filepath}")
                return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error validando archivo {filepath}: {str(e)}")
            return False
    
    def validate_folder(self, folder_path: str) -> bool:
        """Valida que la carpeta exista y sea accesible."""
        try:
            path = Path(folder_path)
            return path.exists() and path.is_dir()
        except Exception as e:
            self.logger.error(f"Error validando carpeta {folder_path}: {str(e)}")
            return False
    
    def get_file_type(self, filepath: str) -> Optional[FileType]:
        """Obtiene el tipo de archivo."""
        ext = Path(filepath).suffix.lower()
        for ft in FileType:
            if ft.value == ext:
                return ft
        return None


# ============================================================================
# GESTOR DE LECTURA DE ARCHIVOS
# ============================================================================
class FileReaderManager:
    """Gestiona múltiples lectores de archivos."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.readers: List[FileReader] = [
            ExcelFileReader(logger),
            CSVFileReader(logger)
        ]
    
    def read_file(self, filepath: str, file_type: FileType) -> List[Dict]:
        """Lee un archivo usando el lector apropiado."""
        for reader in self.readers:
            if reader.supports(file_type):
                return reader.read(filepath)
        
        raise InvalidFileException(f"No hay lector para tipo {file_type}")


# ============================================================================
# CLASE PRINCIPAL DE UNIFICACIÓN
# ============================================================================
class ExcelUnifier:
    """Unificador principal de archivos Excel/CSV con POO limpia."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.validator = FileValidator(logger)
        self.reader_manager = FileReaderManager(logger)
        self.progress_callback: Optional[callable] = None
    
    def set_progress_callback(self, callback: callable) -> None:
        """Establece callback para actualizaciones de progreso."""
        self.progress_callback = callback
    
    def _report_progress(self, progress: UnificationProgress) -> None:
        """Reporta progreso si hay callback."""
        if self.progress_callback:
            self.progress_callback(progress)
    
    def collect_files(self, paths: Union[str, List[str]], recursive: bool = True) -> List[str]:
        """Recolecta archivos válidos de rutas o directorios."""
        valid_files: List[str] = []
        input_paths = [paths] if isinstance(paths, str) else paths
        
        for path_str in input_paths:
            path = Path(path_str)
            
            if path.is_file():
                if self.validator.validate_file(str(path)):
                    valid_files.append(str(path))
            
            elif path.is_dir():
                if not self.validator.validate_folder(str(path)):
                    raise InvalidFolderException(f"Carpeta inválida: {path_str}")
                
                pattern = "**/*" if recursive else "*"
                for file_path in path.glob(pattern):
                    if file_path.is_file() and self.validator.validate_file(str(file_path)):
                        valid_files.append(str(file_path))
        
        if not valid_files:
            raise InvalidFileException("No se encontraron archivos válidos")
        
        self.logger.info(f"Se encontraron {len(valid_files)} archivos válidos")
        return sorted(valid_files)
    
    def unify(self, input_paths: Union[str, List[str]], mode: UnificationMode,
              output_path: str, output_name: str = "unificado.xlsx") -> str:
        """Unifica archivos según el modo especificado."""
        try:
            files = self.collect_files(input_paths)
            
            if mode == UnificationMode.ALL_SHEETS_ALL_FILES:
                return self._unify_all_sheets(files, output_path, output_name)
            elif mode == UnificationMode.ALL_FILES_MERGE:
                return self._merge_all_files(files, output_path, output_name)
            elif mode == UnificationMode.BY_FILE:
                return self._unify_by_file(files, output_path, output_name)
            
        except ExcelUnifierException:
            raise
        except Exception as e:
            raise UnificationException(f"Error durante unificación: {str(e)}")
    
    def _unify_all_sheets(self, files: List[str], output_path: str, output_name: str) -> str:
        """Unifica todas las hojas de todos los archivos manteniendo estructura."""
        all_sheets: Dict[str, pd.DataFrame] = {}
        total_files = len(files)
        
        for file_idx, filepath in enumerate(files):
            try:
                file_type = self.validator.get_file_type(filepath)
                sheet_data = self.reader_manager.read_file(filepath, file_type)
                
                for sheet_info in sheet_data:
                    sheet_name = f"{Path(filepath).stem}_{sheet_info['sheet_name']}"
                    sheet_name = sheet_name[:31]
                    all_sheets[sheet_name] = sheet_info['data']
                
                progress = UnificationProgress(
                    current_file=file_idx + 1,
                    total_files=total_files,
                    percentage=(file_idx + 1) / total_files * 100,
                    status=f"Procesando: {Path(filepath).name}"
                )
                self._report_progress(progress)
                self.logger.info(f"Archivo procesado: {filepath}")
                
            except Exception as e:
                self.logger.error(f"Error procesando {filepath}: {str(e)}")
                continue
        
        return self._save_workbook(all_sheets, output_path, output_name)
    
    def _merge_all_files(self, files: List[str], output_path: str, output_name: str) -> str:
        """Consolida todos los archivos en una sola hoja."""
        consolidated_data: List[pd.DataFrame] = []
        total_files = len(files)
        
        for file_idx, filepath in enumerate(files):
            try:
                file_type = self.validator.get_file_type(filepath)
                sheet_data = self.reader_manager.read_file(filepath, file_type)
                
                for sheet_info in sheet_data:
                    consolidated_data.append(sheet_info['data'])
                
                progress = UnificationProgress(
                    current_file=file_idx + 1,
                    total_files=total_files,
                    percentage=(file_idx + 1) / total_files * 100,
                    status=f"Consolidando: {Path(filepath).name}"
                )
                self._report_progress(progress)
                
            except Exception as e:
                self.logger.error(f"Error consolidando {filepath}: {str(e)}")
                continue
        
        merged_df = pd.concat(consolidated_data, ignore_index=True, sort=False)
        all_sheets = {"Consolidado": merged_df}
        
        return self._save_workbook(all_sheets, output_path, output_name)
    
    def _unify_by_file(self, files: List[str], output_path: str, output_name: str) -> str:
        """Unifica hojas dentro de cada archivo manteniendo separación."""
        all_sheets: Dict[str, pd.DataFrame] = {}
        total_files = len(files)
        
        for file_idx, filepath in enumerate(files):
            try:
                file_type = self.validator.get_file_type(filepath)
                sheet_data = self.reader_manager.read_file(filepath, file_type)
                
                if len(sheet_data) > 1:
                    unified_sheets = [s['data'] for s in sheet_data]
                    merged = pd.concat(unified_sheets, ignore_index=True, sort=False)
                    sheet_name = Path(filepath).stem[:31]
                    all_sheets[sheet_name] = merged
                else:
                    sheet_name = Path(filepath).stem[:31]
                    all_sheets[sheet_name] = sheet_data[0]['data']
                
                progress = UnificationProgress(
                    current_file=file_idx + 1,
                    total_files=total_files,
                    percentage=(file_idx + 1) / total_files * 100,
                    status=f"Unificando: {Path(filepath).name} ({((file_idx + 1) / total_files * 100)} % completado)"
                )
                self._report_progress(progress)
                
            except Exception as e:
                self.logger.error(f"Error procesando {filepath}: {str(e)}")
                continue
        
        return self._save_workbook(all_sheets, output_path, output_name)
    
    def _save_workbook(self, sheets: Dict[str, pd.DataFrame], output_path: str, output_name: str) -> str:
        """Guarda el workbook en formato Excel."""
        try:
            if not sheets:
                raise UnificationException("No hay datos para guardar")
            
            output_dir = Path(output_path)
            if not output_dir.exists():
                output_dir.mkdir(parents=True, exist_ok=True)
            
            output_file = output_dir / output_name
            
            with pd.ExcelWriter(str(output_file), engine='openpyxl') as writer:
                for sheet_name, df in sheets.items():
                    # Limpiar NaN y guardar
                    df.fillna('').to_excel(writer, sheet_name=sheet_name, index=False)
            
            self.logger.info(f"Archivo guardado exitosamente: {output_file}")
            return str(output_file)
            
        except Exception as e:
            raise UnificationException(f"Error guardando archivo: {str(e)}")


# ============================================================================
# INTERFAZ GRÁFICA
# ============================================================================
class ExcelUnifierGUI:
    """Interfaz gráfica moderna con Tkinter."""
    
    def __init__(self, root: Tk):
        self.root = root
        self.logger = LoggerConfigurator.setup()
        self.unifier = ExcelUnifier(self.logger)
        
        self.selected_files: List[str] = []
        self.destination_path: Optional[str] = None
        
        self._setup_window()
        self._build_ui()
    
    def _setup_window(self):
        """Configura la ventana principal."""
        self.root.title("EGROJ182")
        self.root.geometry("1100x750")
        self.root.minsize(1000, 650)
        self.root.resizable(True, True)
        self.root.configure(bg='#f0f0f0')
        
        style = ttk.Style()
        style.theme_use('clam')
    
    def _build_ui(self):
        """Construye la interfaz."""
        # Header
        header_frame = Frame(self.root, bg='#2c3e50')
        header_frame.pack(fill='x', side='top')
        
        title_label = Label(header_frame, text="Excel Unifier Pro - EGROJ182",
                          font=('Segoe UI', 18, 'bold'), bg='#2c3e50', fg='white')
        title_label.pack(pady=3)
        
        subtitle_label = Label(header_frame, text="Unifica múltiples archivos Excel/CSV con opciones avanzadas",
                             font=('Segoe UI', 10), bg='#2c3e50', fg='#ecf0f1')
        subtitle_label.pack(pady=(0, 2))
        
        # Main content frame
        main_frame = Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=10, pady=10, side='top')
        
        self._build_files_section(main_frame)
        self._build_destination_section(main_frame)
        self._build_mode_section(main_frame)
        
        # Frame para Progreso + Botones (lado a lado)
        bottom_frame = Frame(main_frame, bg='#f0f0f0')
        bottom_frame.pack(fill='both', expand=False, padx=5, pady=10)
        
        # Progreso a la izquierda (70%)
        progress_container = Frame(bottom_frame, bg='#f0f0f0')
        progress_container.pack(side='left', fill='both', expand=True, padx=(0, 5))
        self._build_progress_section(progress_container)
        
        # Botones a la derecha (30%)
        buttons_container = Frame(bottom_frame, bg='#f0f0f0')
        buttons_container.pack(side='right', fill='both', expand=False, padx=(5, 0))
        self._build_action_buttons(buttons_container)
    
    def _build_files_section(self, parent: Frame):
        """Construye sección de selección de archivos."""
        files_frame = ttk.LabelFrame(parent, text="Archivos a Unificar", padding=10)
        files_frame.pack(fill='both', expand=True, pady=(0, 10), padx=5)
        
        drop_frame = Frame(files_frame, bg='#ecf0f1', relief='groove', bd=2)
        drop_frame.pack(fill='x', expand=False, pady=10, ipady=30)
        
        drop_label = Label(drop_frame, text="📁 Arrastra archivos aquí o usa los botones",
                         font=('Segoe UI', 11), bg='#ecf0f1', fg='#7f8c8d')
        drop_label.pack(expand=True)
        
        try:
            drop_frame.drop_target_register(DND_FILES)
            drop_frame.dnd_bind('<<Drop>>', self._on_files_drop)
        except Exception as e:
            self.logger.warning(f"Drag and drop no disponible: {e}")
        
        buttons_frame = Frame(files_frame, bg='#f0f0f0')
        buttons_frame.pack(fill='x', pady=10)
        
        ttk.Button(buttons_frame, text="📂 Seleccionar Carpeta", 
                  command=self._select_folder).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(buttons_frame, text="📄 Seleccionar Archivos",
                  command=self._select_files).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(buttons_frame, text="🗑️ Limpiar",
                  command=self._clear_files).pack(side='left', padx=5, fill='x', expand=True)
        
        # Frame para Text con Scrollbar
        text_frame = Frame(files_frame, bg='white')
        text_frame.pack(fill='both', expand=True, pady=(10, 0))
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.files_text = Text(text_frame, height=6, bg='white', font=('Courier', 9),
                              yscrollcommand=scrollbar.set, wrap='word')
        self.files_text.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.files_text.yview)
    
    def _build_destination_section(self, parent: Frame):
        """Construye sección de destino."""
        dest_frame = ttk.LabelFrame(parent, text="Carpeta de Destino", padding=10)
        dest_frame.pack(fill='x', pady=10, padx=5)
        
        button_frame = Frame(dest_frame, bg='#f0f0f0')
        button_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Button(button_frame, text="📍 Seleccionar Destino",
                  command=self._select_destination).pack(side='left', padx=5, fill='x', expand=True)
        
        self.destination_label = Label(dest_frame, text="No seleccionada",
                                      font=('Segoe UI', 9), fg='#7f8c8d', justify='left')
        self.destination_label.pack(fill='x', pady=10)
    
    def _build_mode_section(self, parent: Frame):
        """Construye sección de modo unificación."""
        mode_frame = ttk.LabelFrame(parent, text="Modo de Unificación", padding=1)
        mode_frame.pack(fill='x', pady=1, padx=1)
        
        self.mode_var = StringVar(value="all_sheets")
        
        for mode, desc in [
            ("all_sheets", "Todas las hojas de todos los archivos (mantener estructura)"),
            ("merge", "Consolidar todos los archivos en una sola hoja"),
            ("by_file", "Unificar por archivo (una hoja por archivo)")
        ]:
            ttk.Radiobutton(mode_frame, text=desc, variable=self.mode_var,
                          value=mode).pack(side='left', padx=1, fill='x', expand=True)

        name_frame = Frame(mode_frame, bg='#f0f0f0')
        name_frame.pack(fill='x', pady=15, padx=1)
        
        Label(name_frame, text="Nombre archivo salida:",
             font=('Segoe UI', 9), bg='#f0f0f0').pack(side='left', padx=1)
        
        self.output_name = StringVar(value="unificado.xlsx")
        output_entry = ttk.Entry(name_frame, textvariable=self.output_name)
        output_entry.pack(side='left', padx=5, fill='x', expand=True)
    
    def _build_progress_section(self, parent: Frame):
        """Construye sección de progreso."""
        progress_frame = ttk.LabelFrame(parent, text="Progreso", padding=10)
        progress_frame.pack(fill='both', expand=False, pady=0, padx=0)
        
        self.progress_var = StringVar(value="0")
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                           maximum=100, mode='determinate')
        self.progress_bar.pack(fill='x', pady=5)
        
        self.status_label = Label(progress_frame, text="Esperando...",
                                 font=('Segoe UI', 9), fg='#2c3e50', justify='left')
        self.status_label.pack(anchor='w', pady=(5, 0))
    
    def _build_action_buttons(self, parent: Frame):
        """Construye botones de acción."""
        action_frame = Frame(parent, bg='#f0f0f0')
        action_frame.pack(fill='both', expand=False, pady=20)
        
        btn_unify = ttk.Button(action_frame, text="▶ UNIFICAR ARCHIVOS", 
                              command=self._start_unification)
        btn_unify.pack(side='left', padx=1, pady=1, fill='both', expand=False, ipadx=1, ipady=1)
        
        btn_exit = ttk.Button(action_frame, text="❌ Salir", 
                             command=self.root.quit)
        btn_exit.pack(side='left', padx=1, pady=1, fill='both', ipadx=1, ipady=1)
    
    def _on_files_drop(self, event):
        """Maneja archivos arrastrados."""
        try:
            files = self.root.tk.splitlist(event.data)
            for f in files:
                f = f.strip('{}').strip()
                if os.path.exists(f):
                    if os.path.isfile(f):
                        self.selected_files.append(f)
                    elif os.path.isdir(f):
                        # Si es carpeta, agregar archivos de la carpeta
                        try:
                            collected = self.unifier.collect_files(f, recursive=True)
                            self.selected_files.extend(collected)
                        except:
                            pass
            
            self.selected_files = list(set(self.selected_files))
            self._update_files_display()
        except Exception as e:
            self.logger.error(f"Error en drag and drop: {e}")
    
    def _select_folder(self):
        """Selecciona una carpeta."""
        folder = filedialog.askdirectory(title="Seleccionar carpeta con archivos")
        if folder:
            try:
                files = self.unifier.collect_files(folder, recursive=True)
                self.selected_files = files
                self._update_files_display()
            except Exception as e:
                messagebox.showerror("Error", str(e))
    
    def _select_files(self):
        """Selecciona archivos individuales."""
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All", "*.*")]
        )
        if files:
            self.selected_files.extend(files)
            self.selected_files = list(set(self.selected_files))
            self._update_files_display()
    
    def _clear_files(self):
        """Limpia la lista de archivos."""
        self.selected_files = []
        self._update_files_display()
    
    def _update_files_display(self):
        """Actualiza la visualización de archivos."""
        self.files_text.config(state='normal')
        self.files_text.delete('1.0', 'end')
        
        for f in self.selected_files:
            self.files_text.insert('end', f + '\n')
        
        self.files_text.config(state='disabled')
    
    def _select_destination(self):
        """Selecciona carpeta de destino."""
        folder = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if folder:
            self.destination_path = folder
            self.destination_label.config(text=folder, fg='#27ae60')
    
    def _update_progress(self, progress: UnificationProgress):
        """Actualiza barra de progreso."""
        self.progress_var.set(str(int(progress.percentage)))
        status = f"{progress.status} ({progress.current_file}/{progress.total_files})"
        self.status_label.config(text=status)
        self.root.update_idletasks()
    
    def _start_unification(self):
        """Inicia el proceso de unificación en thread separado."""
        if not self.selected_files:
            messagebox.showwarning("Advertencia", "Selecciona archivos primero")
            return
        
        if not self.destination_path:
            messagebox.showwarning("Advertencia", "Selecciona carpeta de destino")
            return
        
        mode_map = {
            "all_sheets": UnificationMode.ALL_SHEETS_ALL_FILES,
            "merge": UnificationMode.ALL_FILES_MERGE,
            "by_file": UnificationMode.BY_FILE
        }
        mode = mode_map[self.mode_var.get()]
        
        def unify_thread():
            try:
                self.unifier.set_progress_callback(self._update_progress)
                output_file = self.unifier.unify(
                    input_paths=self.selected_files,
                    mode=mode,
                    output_path=self.destination_path,
                    output_name=self.output_name.get()
                )
                
                self.root.after(0, lambda: messagebox.showinfo(
                    "Éxito", f"Unificación completada:\n{output_file}"
                ))
                
                self.root.after(0, self._clear_files)
                self.root.after(0, lambda: self.progress_var.set("0"))
                self.root.after(0, lambda: self.status_label.config(text="Completado"))
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
                self.root.after(0, lambda: self.progress_var.set("0"))
                self.root.after(0, lambda: self.status_label.config(text="Error durante unificación"))
        
        try:
            thread = threading.Thread(target=unify_thread, daemon=True)
            thread.start()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo iniciar el proceso: {str(e)}")


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================
def main():
    """Punto de entrada de la aplicación."""
    try:
        root = TkinterDnD.Tk()  # ← Esto inicializa el soporte DnD
        app = ExcelUnifierGUI(root)
        root.mainloop()
    except Exception as e:
        logging.error(f"Error fatal: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    main()