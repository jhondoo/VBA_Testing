import tkinter as tk
from tkinter import ttk


class SettingsPage(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        # ===== Conteneur principal centré avec marges =====
        main_container = ttk.Frame(self)
        main_container.pack(expand=True, fill="both", padx=40, pady=20)

        # ===== Affichage vertical principal =====
        vertical_layout = ttk.Frame(main_container)
        vertical_layout.pack(expand=True, fill="both")

        # Sections (25% / 25% / 50%)
        section1 = ttk.Frame(vertical_layout)
        section2 = ttk.Frame(vertical_layout)
        section3 = ttk.Frame(vertical_layout)

        section1.pack(expand=True, fill="both", pady=10)
        section2.pack(expand=True, fill="both", pady=10)
        section3.pack(expand=True, fill="both", pady=10)

        self.build_section_1(section1)
        self.build_section_2(section2)
        self.build_section_3(section3)

    # =====================================================
    # =================== SECTION 1 =======================
    # =====================================================
    def build_section_1(self, parent):
        container = self.create_horizontal_section(parent)

        content = container["content"]

        var_section1 = tk.StringVar(value="1")

        for i in range(1, 7):
            rb = ttk.Radiobutton(
                content,
                text=f"lb {i}",
                value=str(i),
                variable=var_section1
            )
            rb.pack(anchor="w", pady=2)

    # =====================================================
    # =================== SECTION 2 =======================
    # =====================================================
    def build_section_2(self, parent):
        container = self.create_horizontal_section(parent)
        content = container["content"]

        for _ in range(3):
            sub = ttk.Frame(content)
            sub.pack(fill="x", pady=5)

            var = tk.StringVar(value="A")

            ttk.Radiobutton(
                sub, text="label temp", value="A", variable=var
            ).pack(anchor="w")

            ttk.Radiobutton(
                sub, text="label temp", value="B", variable=var
            ).pack(anchor="w")

    # =====================================================
    # =================== SECTION 3 =======================
    # =====================================================
    def build_section_3(self, parent):
        container = self.create_horizontal_section(parent)
        content = container["content"]

        for _ in range(4):
            sub = ttk.Frame(content)
            sub.pack(fill="x", pady=8)

            var = tk.StringVar(value="1")

            for i in range(1, 4):
                ttk.Radiobutton(
                    sub,
                    text="label temp",
                    value=str(i),
                    variable=var
                ).pack(side="left", padx=10)

    # =====================================================
    # ========== SECTION HORIZONTALE GENERIQUE ============
    # =====================================================
    def create_horizontal_section(self, parent):
        # Conteneur avec padding (design moderne)
        section = ttk.Frame(parent, padding=15)
        section.pack(expand=True, fill="both")

        section.columnconfigure(0, weight=1)
        section.columnconfigure(1, weight=0)
        section.columnconfigure(2, weight=3)

        # Bloc image (25%)
        image_placeholder = ttk.Frame(section, relief="ridge")
        image_placeholder.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        ttk.Label(image_placeholder, text="Image").place(
            relx=0.5, rely=0.5, anchor="center"
        )

        # Séparateur vertical
        separator = ttk.Separator(section, orient="vertical")
        separator.grid(row=0, column=1, sticky="ns")

        # Contenu (75%)
        content = ttk.Frame(section)
        content.grid(row=0, column=2, sticky="nsew", padx=(10, 0))

        return {
            "section": section,
            "image": image_placeholder,
            "content": content
        }
