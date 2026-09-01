import os
import threading
from tkinter import filedialog, messagebox
import customtkinter as ctk

# Dönüştürme kütüphaneleri
from pdf2docx import Converter
from docx2pdf import convert as convert_docx_to_pdf

# Karanlık Mod Ayarları
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class UniversalConverterApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        # Pencere Boyutları
        self.title("PDF & Word Dönüştürücü")
        self.geometry("620x580")
        self.resizable(False, False)

        self.selected_file_path = None
        self.output_directory = None  # Özel Hedef Klasör Yolu
        self.conversion_mode = None  # "pdf2docx" veya "docx2pdf"

        # Ana Çerçeve
        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.pack(padx=30, pady=25, fill="both", expand=True)

        # Başlık
        self.label_title = ctk.CTkLabel(
            self.main_frame,
            text="PDF ⇄ Word Dönüştürücü",
            font=ctk.CTkFont(size=26, weight="bold"),
        )
        self.label_title.pack(pady=(25, 15))

        # 1. Dosya Seçme Butonu
        self.btn_select = ctk.CTkButton(
            self.main_frame,
            text="📁 PDF veya Word Dosyası Seç",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.select_file,
            width=280,
            height=42,
            corner_radius=10,
        )
        self.btn_select.pack(pady=8)

        # Seçilen Dosya Bilgisi Etiketi
        self.label_file = ctk.CTkLabel(
            self.main_frame,
            text="Henüz bir dosya seçilmedi",
            font=ctk.CTkFont(size=13, slant="italic"),
            text_color="#A0A0A0",
            wraplength=500,
        )
        self.label_file.pack(pady=8)

        # 2. Kayıt Klasörü Seçme Butonu
        self.btn_select_dir = ctk.CTkButton(
            self.main_frame,
            text="📂 Kayıt Klasörünü Değiştir",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.select_output_directory,
            fg_color="#4A5568",
            hover_color="#2D3748",
            width=280,
            height=38,
            corner_radius=10,
        )
        self.btn_select_dir.pack(pady=8)

        # Hedef Klasör Bilgisi Etiketi
        self.label_dir = ctk.CTkLabel(
            self.main_frame,
            text="Kayıt Yeri: Dosyanın Bulunduğu Klasör (Varsayılan)",
            font=ctk.CTkFont(size=12),
            text_color="#888888",
            wraplength=500,
        )
        self.label_dir.pack(pady=5)

        # 3. Dönüştürme Butonu
        self.btn_convert = ctk.CTkButton(
            self.main_frame,
            text="🔄 Dönüştür",
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self.start_conversion_thread,
            fg_color="#27ae60",
            hover_color="#1e8449",
            width=280,
            height=45,
            corner_radius=10,
            state="disabled",
        )
        self.btn_convert.pack(pady=15)

        # İlerleme Çubuğu
        self.progress_bar = ctk.CTkProgressBar(self.main_frame, width=420, height=12)
        self.progress_bar.pack(pady=(10, 5))
        self.progress_bar.set(0)

        # Durum Metni
        self.label_status = ctk.CTkLabel(
            self.main_frame,
            text="",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.label_status.pack(pady=5)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Dosya Seçin",
            filetypes=[
                ("Tüm Desteklenenler", "*.pdf;*.docx"),
                ("PDF Dosyaları", "*.pdf"),
                ("Word Dosyaları", "*.docx"),
            ],
        )
        if file_path:
            self.selected_file_path = file_path
            dosya_adi = os.path.basename(file_path)
            uzanti = os.path.splitext(file_path)[1].lower()

            if uzanti == ".pdf":
                self.conversion_mode = "pdf2docx"
                islem_metni = "İşlem: PDF ➔ Word"
            elif uzanti == ".docx":
                self.conversion_mode = "docx2pdf"
                islem_metni = "İşlem: Word ➔ PDF"
            else:
                return

            self.label_file.configure(
                text=f"{dosya_adi} ({islem_metni})",
                text_color="#3498db",
                font=ctk.CTkFont(size=14, weight="bold"),
            )
            self.btn_convert.configure(state="normal")
            self.label_status.configure(text="")
            self.progress_bar.set(0)

    def select_output_directory(self):
        selected_dir = filedialog.askdirectory(title="Kayıt Klasörünü Seçin")
        if selected_dir:
            self.output_directory = selected_dir
            self.label_dir.configure(
                text=f"Kayıt Yeri: {selected_dir}",
                text_color="#2ecc71",
            )

    def start_conversion_thread(self):
        threading.Thread(target=self.convert_file, daemon=True).start()

    def convert_file(self):
        if not self.selected_file_path or not self.conversion_mode:
            return

        try:
            self.btn_select.configure(state="disabled")
            self.btn_select_dir.configure(state="disabled")
            self.btn_convert.configure(state="disabled")
            self.label_status.configure(
                text="Dönüştürülüyor, lütfen bekleyin...", text_color="#f1c40f"
            )
            self.progress_bar.start()

            # Kayıt yolunu belirle
            dosya_adi_uzantisiz = os.path.splitext(
                os.path.basename(self.selected_file_path)
            )[0]
            
            # Özel klasör seçildiyse oraya, seçilmediyse kaynak dosyanın yanına kaydet
            hedef_klasor = (
                self.output_directory
                if self.output_directory
                else os.path.dirname(self.selected_file_path)
            )

            if self.conversion_mode == "pdf2docx":
                output_path = os.path.join(hedef_klasor, f"{dosya_adi_uzantisiz}.docx")
                cv = Converter(self.selected_file_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()

            elif self.conversion_mode == "docx2pdf":
                output_path = os.path.join(hedef_klasor, f"{dosya_adi_uzantisiz}.pdf")
                convert_docx_to_pdf(self.selected_file_path, output_path)

            self.progress_bar.stop()
            self.progress_bar.set(1.0)
            self.label_status.configure(
                text="Dönüştürme Başarıyla Tamamlandı!", text_color="#2ecc71"
            )

            messagebox.showinfo(
                "Başarılı",
                f"Dosya başarıyla dönüştürüldü:\n\n{output_path}",
            )

        except Exception as e:
            self.progress_bar.stop()
            self.progress_bar.set(0)
            self.label_status.configure(text="Hata Oluştu!", text_color="#e74c3c")
            messagebox.showerror(
                "Hata", f"Dönüştürme sırasında hata oluştu:\n\n{str(e)}"
            )

        finally:
            self.btn_select.configure(state="normal")
            self.btn_select_dir.configure(state="normal")
            self.btn_convert.configure(state="normal")


if __name__ == "__main__":
    app = UniversalConverterApp()
    app.mainloop()