# pyright: reportMissingImports=false
import sys
import os
import tkinter as tk
from tkinter import ttk

from industry_section import IndustrySectorSection

class StockSearchApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("실시간 종목 탐색기 (Stock Search Standalone)")
        self.root.geometry("1300x750")
        self.root.minsize(1200, 650)
        
        # ── Neo-Brutalist Theme ──────────────────────
        self.bg_dark      = "#F5F0E8"   # Warm Off-White
        self.bg_card      = "#FFFFFF"   # Pure White
        self.bg_input     = "#FFFFFF"   # Input field background
        self.text_light   = "#FFFFFF"   # Light text on buttons
        self.text_dark    = "#1A1A1A"   # Basic dark text
        self.text_muted   = "#4A4A4A"   # Subtitle/muted text
        self.primary      = "#1A1A1A"   # Primary Black
        self.accent       = "#FFCC00"   # Vivid Yellow
        self.hover        = "#FFCC00"   
        self.hover_dark   = "#E6B800"   
        self.tertiary     = "#0055FF"   # Bauhaus Blue
        self.tertiary_hov = "#003DD6"   
        self.error        = "#E63B2E"   # Signal Red
        self.gray         = "#4A4A4A"   
        self.border_color = "#1A1A1A"   # Border color
        self.btn_rss      = "#0055FF"   
        # ────────────────────────────────────────────

        self.root.configure(bg=self.bg_dark)
        self.setup_styles()
        self.create_widgets()
        
    def setup_styles(self) -> None:
        self.style = ttk.Style()
        self.style.theme_use("default")

        # Notebook
        self.style.configure("TNotebook",
            background=self.bg_dark, borderwidth=0)
        self.style.configure("TNotebook.Tab",
            background=self.bg_card,
            foreground=self.text_dark,
            padding=[22, 9],
            font=("맑은 고딕", 10, "bold"),
            borderwidth=2,
            lightcolor=self.border_color,
            darkcolor=self.border_color,
        )
        self.style.map("TNotebook.Tab",
            background=[("selected", self.accent)],
            foreground=[("selected", self.text_dark)],
        )

        # Scrollbar
        self.style.configure("Vertical.TScrollbar",
            background=self.gray,
            troughcolor=self.bg_dark,
            arrowcolor=self.text_light,
            borderwidth=0,
        )

        # Combobox
        self.style.configure("TCombobox",
            fieldbackground=self.bg_input,
            background=self.primary,
            foreground=self.text_dark,
            arrowcolor=self.text_dark,
            selectbackground=self.accent,
            selectforeground=self.text_dark,
            borderwidth=2,
        )
        self.style.map("TCombobox",
            fieldbackground=[("readonly", self.bg_input)],
            foreground=[("readonly", self.text_dark)],
            selectbackground=[("readonly", self.accent)],
        )

        # Dropdown Listbox elements
        self.root.option_add("*TCombobox*Listbox.background",       self.bg_card)
        self.root.option_add("*TCombobox*Listbox.foreground",       self.text_dark)
        self.root.option_add("*TCombobox*Listbox.selectBackground", self.accent)
        self.root.option_add("*TCombobox*Listbox.selectForeground", self.text_dark)
        self.root.option_add("*TCombobox*Listbox.font",             "맑은고딕 10")
        
        # Configure global treeview style
        self.style.configure("Treeview",
            font=("맑은 고딕", 9),
            background=self.bg_card,
            fieldbackground=self.bg_card,
            foreground=self.text_dark,
            rowheight=24,
            borderwidth=1,
        )
        self.style.configure("Treeview.Heading",
            font=("맑은 고딕", 9, "bold"),
            background=self.bg_dark,
            foreground=self.text_dark,
        )
        
    def _theme_dict(self) -> dict:
        return {
            "bg_card":      self.bg_card,
            "bg_input":     self.bg_input,
            "text_light":   self.text_light,
            "text_dark":    self.text_dark,
            "text_muted":   self.text_muted,
            "accent":       self.accent,
            "hover":        self.hover,
            "hover_dark":   self.hover_dark,
            "btn_rss":      self.btn_rss,
            "tertiary":     self.tertiary,
            "tertiary_hov": self.tertiary_hov,
            "primary":      self.primary,
            "border_color": self.border_color,
            "bg_dark":      self.bg_dark,
        }

    def create_widgets(self) -> None:
        # Header (Bauhaus Black Bar)
        header = tk.Frame(self.root, bg=self.primary, pady=12, bd=2, relief="solid")
        header.pack(fill="x", side="top")

        tk.Label(
            header,
            text="STOCK SEARCH EXPLORER",
            font=("맑은 고딕", 16, "bold"),
            bg=self.primary, fg=self.text_light,
        ).pack(side="left", padx=20)

        tk.Label(
            header,
            text="실시간 종목 탐색 & 매매 동향 분석",
            font=("맑은 고딕", 9, "italic"),
            bg=self.primary, fg=self.accent,
        ).pack(side="left", padx=4)

        # Main Container Frame
        self.main_container = tk.Frame(self.root, bg=self.bg_dark)
        self.main_container.pack(fill="both", expand=True, padx=16, pady=15)

        # Instantiate and build Industry Sector Section
        self.industry_section = IndustrySectorSection(
            self.root, self._theme_dict()
        )
        self.industry_section.build(self.main_container)

def main():
    root = tk.Tk()
    
    # Simple icon fallback or styling
    try:
        # Set app window icon (if available)
        pass
    except Exception:
        pass
        
    app = StockSearchApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
