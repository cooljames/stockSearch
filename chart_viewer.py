# pyright: reportMissingImports=false
import re
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

import matplotlib
matplotlib.use("TkAgg")  # Use TkAgg backend for embedding in Tkinter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import yfinance as yf

# Font settings for Korean character support
matplotlib.rcParams["font.family"] = "Malgun Gothic"
matplotlib.rcParams["axes.unicode_minus"] = False

PERIOD_MAP = {
    "1주": "7d",
    "1개월": "1mo",
    "3개월": "3mo",
    "6개월": "6mo",
    "1년": "1y",
    "3년": "3y",
    "5년": "5y"
}

class ChartViewer(tk.Toplevel):
    def __init__(self, parent, stocks: list, theme: dict):
        super().__init__(parent)
        self.stocks = [(sym.strip().upper(), name.strip()) for sym, name in stocks]
        self.theme = theme
        
        # Display title based on selection count
        if len(self.stocks) == 1:
            title_text = f"{self.stocks[0][1]} ({self.stocks[0][0]}) - 실시간 차트 및 종목 정보"
        else:
            title_text = f"{self.stocks[0][1]} 외 {len(self.stocks)-1}개 종목 비교 - 실시간 차트"
            
        self.title(title_text)
        self.geometry("850x700")
        self.minsize(750, 600)
        self.configure(bg=theme.get("bg_dark", "#F5F0E8"))
        
        # Theme variables
        self.bg_dark = theme.get("bg_dark", "#F5F0E8")
        self.bg_card = theme.get("bg_card", "#FFFFFF")
        self.bg_input = theme.get("bg_input", "#FFFFFF")
        self.text_light = theme.get("text_light", "#FFFFFF")
        self.text_dark = theme.get("text_dark", "#1A1A1A")
        self.text_muted = theme.get("text_muted", "#4A4A4A")
        self.primary = theme.get("primary", "#1A1A1A")
        self.accent = theme.get("accent", "#FFCC00")
        self.hover_dark = theme.get("hover_dark", "#E6B800")
        self.border_color = theme.get("border_color", "#1A1A1A")
        
        self.current_period = "3개월"
        self.hist_data_dict = {}  # {symbol: hist_data}
        self.ticker_info_dict = {}  # {symbol: fast_info}
        self.loaded_stocks = []  # successfully loaded (symbol, name)
        
        self._build_ui()
        self._load_data_thread()
        
    def _build_ui(self):
        # Header block
        header = tk.Frame(self, bg=self.primary, pady=12, bd=2, relief="solid")
        header.pack(fill="x", side="top", padx=10, pady=(10, 5))
        
        if len(self.stocks) == 1:
            header_text = f"  {self.stocks[0][1]}  ({self.stocks[0][0]})  "
        else:
            header_text = f"  {self.stocks[0][1]} 외 {len(self.stocks)-1}개 종목 비교  "
            
        tk.Label(
            header, text=header_text, font=("맑은 고딕", 14, "bold"),
            bg=self.primary, fg=self.text_light
        ).pack(side="left", padx=10)
        
        # Info Panel
        self.info_frame = tk.Frame(self, bg=self.bg_card, bd=2, relief="solid", pady=5)
        self.info_frame.pack(fill="x", padx=10, pady=5)
        
        self.price_container = tk.Frame(self.info_frame, bg=self.bg_card)
        self.price_container.pack(fill="x", expand=True, padx=10)
        
        # Main Layout (Chart + Controls)
        self.main_content = tk.Frame(self, bg=self.bg_dark)
        self.main_content.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Left side controls
        ctrl_frame = tk.Frame(self.main_content, bg=self.bg_card, bd=2, relief="solid", width=160)
        ctrl_frame.pack(side="left", fill="y", padx=(0, 5))
        ctrl_frame.pack_propagate(False)
        
        tk.Label(
            ctrl_frame, text="조회 기간", font=("맑은 고딕", 9, "bold"),
            bg=self.bg_card, fg=self.text_dark
        ).pack(pady=(15, 8))
        
        self.period_var = tk.StringVar(value=self.current_period)
        for p in PERIOD_MAP.keys():
            rbtn = tk.Radiobutton(
                ctrl_frame, text=p, variable=self.period_var, value=p,
                font=("맑은 고딕", 9), bg=self.bg_card, fg=self.text_dark,
                activebackground=self.bg_card, selectcolor=self.bg_input,
                command=self._on_period_change
            )
            rbtn.pack(anchor="w", padx=25, pady=4)
            
        # Export Button
        export_btn = tk.Button(
            ctrl_frame, text="💾 차트 이미지 저장", font=("맑은 고딕", 9, "bold"),
            bg=self.primary, fg=self.text_light, bd=1, relief="solid",
            cursor="hand2", padx=10, pady=6, command=self._save_chart_image
        )
        export_btn.pack(side="bottom", fill="x", padx=10, pady=15)
        export_btn.bind("<Enter>", lambda e: e.widget.configure(bg=self.accent, fg=self.text_dark))
        export_btn.bind("<Leave>", lambda e: e.widget.configure(bg=self.primary, fg=self.text_light))
        
        # Right side Chart canvas
        self.chart_container = tk.Frame(self.main_content, bg=self.bg_card, bd=2, relief="solid")
        self.chart_container.pack(side="right", fill="both", expand=True)
        
        self.loading_lbl = tk.Label(
            self.chart_container, text="yfinance 시세 데이터를 조회 중입니다...\n잠시만 기다려 주세요.",
            font=("맑은 고딕", 10), bg=self.bg_card, fg=self.text_muted
        )
        self.loading_lbl.pack(expand=True)
        
        self.fig, self.ax = plt.subplots(figsize=(6, 4.5), dpi=100)
        self.fig.patch.set_facecolor(self.bg_card)
        self.ax.set_facecolor(self.bg_dark)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_container)
        
    def _load_data_thread(self):
        self.loading_lbl.pack(expand=True)
        self.canvas.get_tk_widget().pack_forget()
        threading.Thread(target=self._load_data, daemon=True).start()
        
    def _load_data(self):
        try:
            from concurrent.futures import ThreadPoolExecutor
            
            self.hist_data_dict.clear()
            self.ticker_info_dict.clear()
            self.loaded_stocks.clear()
            
            yf_period = PERIOD_MAP[self.current_period]
            
            def fetch_single(item):
                symbol, name = item
                symbol_to_fetch = symbol
                if re.match(r"^\d{6}$", symbol_to_fetch):
                    symbol_to_fetch += ".KS"
                
                try:
                    tk_obj = yf.Ticker(symbol_to_fetch)
                    hist = tk_obj.history(period=yf_period, auto_adjust=False)
                    if hist is not None and not hist.empty:
                        fast_info = None
                        try:
                            fast_info = tk_obj.fast_info
                        except Exception:
                            pass
                        return symbol, name, hist, fast_info
                except Exception as e:
                    print(f"Error fetching {symbol}: {e}")
                return symbol, name, None, None

            with ThreadPoolExecutor(max_workers=5) as executor:
                results = list(executor.map(fetch_single, self.stocks))
                
            for symbol, name, hist, fast_info in results:
                if hist is not None:
                    self.hist_data_dict[symbol] = hist
                    if fast_info:
                        self.ticker_info_dict[symbol] = fast_info
                    self.loaded_stocks.append((symbol, name))
                    
            self.after(0, self._on_data_loaded)
        except Exception as e:
            self.after(0, lambda: self._on_data_error(str(e)))
            
    def _on_data_loaded(self):
        if not self.loaded_stocks:
            self._on_data_error("선택한 모든 종목의 시세 데이터를 불러오지 못했습니다.")
            return
            
        self.loading_lbl.pack_forget()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        
        self._update_info_panel()
        self._draw_chart()
        
    def _on_data_error(self, err_msg: str):
        self.loading_lbl.configure(text=f"❌ 데이터를 불러오는데 실패했습니다.\n\n오류 내용:\n{err_msg}")
        messagebox.showerror("데이터 조회 오류", f"시세 데이터를 가져오지 못했습니다:\n{err_msg}")
        
    def _update_info_panel(self):
        # Clear existing widgets in price_container
        for w in self.price_container.winfo_children():
            w.destroy()
            
        for idx, (symbol, name) in enumerate(self.loaded_stocks):
            hist = self.hist_data_dict[symbol]
            close_series = hist["Close"].dropna()
            if close_series.empty:
                continue
                
            last_val = float(close_series.iloc[-1])
            prev_val = float(close_series.iloc[-2]) if len(close_series) > 1 else last_val
            
            # Try live price from fast_info if valid
            fast_info = self.ticker_info_dict.get(symbol)
            if fast_info and hasattr(fast_info, "last_price") and fast_info.last_price is not None:
                last_val = float(fast_info.last_price)
                if hasattr(fast_info, "previous_close") and fast_info.previous_close is not None:
                    prev_val = float(fast_info.previous_close)
                    
            chg = last_val - prev_val
            chg_rt = (chg / prev_val) * 100 if prev_val else 0.0
            
            sign = "▲" if chg > 0 else ("▼" if chg < 0 else "−")
            color = "#E63B2E" if chg > 0 else ("#0055FF" if chg < 0 else self.text_dark)
            
            is_kr = ".KS" in symbol or ".KQ" in symbol or re.match(r"^\d{6}", symbol)
            fmt = "{:,.0f}원" if is_kr else "${:,.2f}"
            chg_str = f"{abs(chg):,.0f}" if is_kr else f"{abs(chg):,.2f}"
            
            # Create a small Brutalist card/row for each stock
            stock_row = tk.Frame(self.price_container, bg=self.bg_card, bd=1, relief="solid")
            if len(self.loaded_stocks) <= 2:
                stock_row.pack(side="left", fill="both", expand=True, padx=5, pady=2)
            else:
                r = idx // 3
                c = idx % 3
                stock_row.grid(row=r, column=c, sticky="nsew", padx=5, pady=3)
                self.price_container.grid_columnconfigure(c, weight=1)
                
            # Content inside the row/card
            lbl_name = tk.Label(
                stock_row, text=f"{name} ({symbol})", font=("맑은 고딕", 9, "bold"),
                bg=self.bg_card, fg=self.text_muted
            )
            lbl_name.pack(anchor="w", padx=8, pady=(4, 0))
            
            lbl_price = tk.Label(
                stock_row, text=f"{fmt.format(last_val)}   {sign} {chg_str} ({chg_rt:+.2f}%)",
                font=("맑은 고딕", 11, "bold"), bg=self.bg_card, fg=color
            )
            lbl_price.pack(anchor="w", padx=8, pady=(0, 4))
            
    def _draw_chart(self):
        self.ax.clear()
        
        # Grid styling
        self.ax.grid(True, linestyle="--", alpha=0.5, color=self.text_muted)
        
        LINE_COLORS = ["#1A1A1A", "#0055FF", "#E63B2E", "#00B050", "#7030A0"]
        
        is_multi = len(self.loaded_stocks) > 1
        
        for idx, (symbol, name) in enumerate(self.loaded_stocks):
            hist = self.hist_data_dict[symbol]
            close_data = hist["Close"].dropna()
            if close_data.empty:
                continue
                
            dates = close_data.index
            prices = close_data.values
            
            line_color = LINE_COLORS[idx % len(LINE_COLORS)]
            
            if not is_multi:
                # Single stock: Plot absolute price
                is_up = prices[-1] >= prices[0]
                line_color = "#E63B2E" if is_up else "#0055FF"
                self.ax.plot(dates, prices, color=line_color, linewidth=2, label=name)
                # Fill under the line
                self.ax.fill_between(dates, prices, min(prices)*0.99, facecolor=line_color, alpha=0.1)
                
                title_str = f"{name} ({symbol}) - {self.current_period} 시세 변동"
            else:
                # Multiple stocks: Plot relative return (%)
                initial_price = prices[0]
                if initial_price == 0:
                    continue
                returns = ((prices / initial_price) - 1.0) * 100.0
                self.ax.plot(dates, returns, color=line_color, linewidth=2, label=name)
                
                title_str = f"선택 종목 {self.current_period} 수익률 비교 (%)"
                
        # Set titles & labels
        self.ax.set_title(title_str, fontdict={"fontsize": 11, "fontweight": "bold", "color": self.text_dark})
        
        # Legend
        self.ax.legend(loc="upper left", prop={"size": 8})
        
        # Format X axis dates nicely
        import matplotlib.dates as mdates
        if self.current_period in ("1주", "1개월"):
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))
        else:
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
            
        self.ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        self.fig.autofmt_xdate()
        
        # Y axis formatting
        if not is_multi:
            symbol = self.loaded_stocks[0][0]
            is_kr = ".KS" in symbol or ".KQ" in symbol or re.match(r"^\d{6}", symbol)
            if is_kr:
                self.ax.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
            else:
                self.ax.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, p: f"${x:,.2f}"))
        else:
            self.ax.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, p: f"{x:+.1f}%"))
            
        self.ax.tick_params(colors=self.text_dark, labelsize=9)
        for spine in self.ax.spines.values():
            spine.set_edgecolor(self.border_color)
            spine.set_linewidth(1.5)
            
        self.fig.tight_layout()
        self.canvas.draw()
        
    def _on_period_change(self):
        self.current_period = self.period_var.get()
        self._load_data_thread()
        
    def _save_chart_image(self):
        if not self.loaded_stocks:
            messagebox.showwarning("저장 불가", "차트 데이터가 존재하지 않습니다.")
            return
            
        if len(self.loaded_stocks) == 1:
            default_file = f"{self.loaded_stocks[0][1]}_{self.loaded_stocks[0][0]}_{self.current_period}_차트.png"
        else:
            default_file = f"종목비교_{self.current_period}_차트.png"
            
        path = filedialog.asksaveasfilename(
            title="차트 이미지 저장",
            defaultextension=".png",
            initialfile=default_file,
            filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("All Files", "*.*")]
        )
        if not path:
            return
            
        try:
            self.fig.savefig(path, dpi=300, facecolor=self.bg_card)
            messagebox.showinfo("저장 완료", f"차트 이미지를 성공적으로 저장했습니다:\n\n{path}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"이미지 저장에 실패했습니다:\n{e}")
