import tkinter as tk
from tkinter import ttk, messagebox
import yfinance as yf
import logging
from concurrent.futures import ThreadPoolExecutor
import queue
from functools import lru_cache
import pandas as pd
import os
import numpy as np

# Constants
TOOLTIP_DELAY = 500
DEFAULT_FONT = ('Helvetica', 12)
BOLD_FONT = ('Helvetica', 12, 'bold')
WINDOW_TITLE = "Comprehensive Stock Analysis Tool"
WINDOW_GEOMETRY = "1600x900"
CACHE_SIZE = 128
MAX_WORKERS = 5
MAX_SUGGESTIONS = 10

# Configure logging to write to a file
logging.basicConfig(filename='app.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class ToolTip:
    def __init__(self, widget, text, delay=TOOLTIP_DELAY):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.delay = delay
        self.id = None
        self.widget.bind("<Enter>", self.schedule_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
        self.widget.bind("<FocusOut>", self.hide_tooltip)

    def schedule_tooltip(self, event):
        self.id = self.widget.after(self.delay, self.show_tooltip, event)

    def show_tooltip(self, event):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, font=("tahoma", "10", "normal"))
        label.pack()

    def hide_tooltip(self, event):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

class StockAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_GEOMETRY)
        
        self.data_cache = {}
        self.result_queue = queue.Queue()
        self.executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)
        self.colorblind_mode = False
        self.dark_mode = False
        self.tag_colors = {}

        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.all_tickers = self.get_all_tickers()

        self.setup_style()
        self.setup_gui()

    def get_all_tickers(self):
        return ["AAPL", "MSFT", "GOOGL", "TSLA", "AMZN"]
    
    def get_stock_price_and_change(self, symbol, period='1d'):
        try:
            stock = yf.Ticker(symbol)
            hist = stock.history(period=period)
            if hist.empty:
                return None, None, None
            current_price = hist['Close'].iloc[-1]
            open_price = hist['Open'].iloc[0]
            price_change = current_price - open_price
            price_change_percent = (price_change / open_price) * 100
            return current_price, price_change, price_change_percent
        except Exception as e:
            logging.error(f"Error fetching stock price for {symbol}: {e}")
            return None, None, None

    def wrap_text(self, text, width):
        words = text.split()
        lines = []
        current_line = []
        current_length = 0
        for word in words:
            if current_length + len(word) <= width:
                current_line.append(word)
                current_length += len(word) + 1
            else:
                lines.append(' '.join(current_line))
                current_line = [word]
                current_length = len(word) + 1
        if current_line:
            lines.append(' '.join(current_line))
        return '\n'.join(lines)

    def setup_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Define colors
        if self.dark_mode:
            bg_color = '#2C3E50'  # Dark blue-gray
            fg_color = '#ECF0F1'  # Light gray
            accent_color = '#3498DB'  # Bright blue
            button_color = '#2980B9'  # Darker blue for buttons
            tree_bg = '#34495E'  # Slightly lighter than bg_color for contrast
        else:
            bg_color = '#ECF0F1'  # Light gray
            fg_color = '#2C3E50'  # Dark blue-gray
            accent_color = '#3498DB'  # Bright blue
            button_color = '#2980B9'  # Darker blue for buttons
            tree_bg = '#FFFFFF'  # White background for treeview in light mode
        
        # Configure styles
        style.configure('TFrame', background=bg_color)
        style.configure('TLabel', background=bg_color, foreground=fg_color, font=('Helvetica', 12))
        style.configure('TEntry', fieldbackground=fg_color, foreground=bg_color, font=('Helvetica', 12))
        style.configure('TButton', background=button_color, foreground=fg_color, font=('Helvetica', 12, 'bold'), padding=10)
        style.map('TButton', background=[('active', accent_color)])
        
        # Treeview (results table) styling
        style.configure('Treeview', 
                        background=tree_bg, 
                        foreground=fg_color, 
                        rowheight=100,  # Set a default row height
                        fieldbackground=tree_bg, 
                        font=('Helvetica', 12))
        style.configure('Treeview.Heading', 
                        background=button_color, 
                        foreground=fg_color, 
                        font=('Helvetica', 12, 'bold'))
        style.map('Treeview', background=[('selected', accent_color)], foreground=[('selected', fg_color)])

        # Checkbox styling
        style.configure('TCheckbutton', background=bg_color, foreground=fg_color, font=('Helvetica', 12))
        
        # Combobox styling
        style.configure('TCombobox', fieldbackground=fg_color, foreground=bg_color, font=('Helvetica', 12))

        # Set colors for the main window
        self.root.configure(bg=bg_color)

        # Configure tag colors
        if self.colorblind_mode:
            self.tag_colors = {
                'green': '#377eb8',  # Blue
                'red': '#ff7f00',    # Orange
                'black': fg_color
            }
        else:
            self.tag_colors = {
                'green': '#2ecc71',  # Green
                'red': '#e74c3c',    # Red
                'black': fg_color
            }

    def setup_gui(self):
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=0)
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.root.after(100, self.process_queue)

        # Input frame
        input_frame = ttk.Frame(self.root, padding="10", style='TFrame')
        input_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        input_frame.grid_columnconfigure(1, weight=1)

        ticker_label = ttk.Label(input_frame, text="Enter Ticker Symbol(s):", style='TLabel')
        ticker_label.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.ticker_entry = ttk.Entry(input_frame, style='TEntry', width=50)
        self.ticker_entry.grid(row=0, column=1, sticky="ew")

        ToolTip(ticker_label, "Enter ticker symbols separated by commas. Example: AAPL,MSFT,GOOGL")

        # Period dropdown
        period_label = ttk.Label(input_frame, text="Select Period:", style='TLabel')
        period_label.grid(row=0, column=2, padx=(10, 5), sticky="w")
        self.period_var = tk.StringVar(value="1d")
        self.period_combo = ttk.Combobox(input_frame, textvariable=self.period_var, values=["1d", "5d", "1mo", "3mo", "6mo", "1y", "ytd", "max"], style='TCombobox', width=5)
        self.period_combo.grid(row=0, column=3, padx=(0, 10), sticky="w")

        ToolTip(period_label, "Select the time period for stock price change calculation")

        # Options frame
        options_frame = ttk.Frame(input_frame, style='TFrame')
        options_frame.grid(row=0, column=4, padx=(10, 0), sticky="e")

        self.colorblind_var = tk.BooleanVar()
        colorblind_check = ttk.Checkbutton(options_frame, text="Colorblind Mode", variable=self.colorblind_var, command=self.toggle_colorblind_mode, style='TCheckbutton')
        colorblind_check.grid(row=0, column=0, padx=(0, 10))
        ToolTip(colorblind_check, "Toggle colorblind mode for better accessibility")

        self.dark_mode_var = tk.BooleanVar()
        dark_mode_check = ttk.Checkbutton(options_frame, text="Dark Mode", variable=self.dark_mode_var, command=self.toggle_dark_mode, style='TCheckbutton')
        dark_mode_check.grid(row=0, column=1, padx=(0, 10))
        ToolTip(dark_mode_check, "Toggle dark mode for better accessibility")

        # Button frame
        button_frame = ttk.Frame(self.root, padding="10", style='TFrame')
        button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))

        methods = [
            ("Buffett Method", self.calculate_value, "Calculate stock value using the Buffett method"),
            ("Brandes Method", self.evaluate_stock, "Evaluate stock using the Brandes method"),
            ("Hartz, Millsap, Hill Method", self.calculate_hartz_millsap_hill, "Calculate stock value using the Hartz, Millsap, Hill method"),
            ("Pabrai Method", self.calculate_intrinsic_value, "Calculate intrinsic stock value using the Pabrai method"),
            ("Hempton Nutty Method", self.calculate_hempton_nutty, "Evaluate stock using the Hempton Nutty method")
        ]

        self.buttons = {}
        for i, (method_name, method_function, tooltip_text) in enumerate(methods):
            self.buttons[method_name] = ttk.Button(button_frame, text=method_name,
                                                command=lambda f=method_function: self.run_analysis_method(f),
                                                style='TButton')
            self.buttons[method_name].grid(row=0, column=i, padx=5, pady=5)
            ToolTip(self.buttons[method_name], tooltip_text)

        run_all_button = ttk.Button(button_frame, text="Run All", command=self.run_all_methods, style='TButton')
        run_all_button.grid(row=0, column=len(methods), padx=5, pady=5)
        ToolTip(run_all_button, "Run all analysis methods")

        # Results frame
        result_frame = ttk.Frame(self.root, padding="10", style='TFrame')
        result_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)

        columns = ["Ticker", "Buffett", "Brandes", "Hartz, Millsap, Hill", "Pabrai", "Hempton Nutty", "Current Price"]
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show='headings', style='Treeview')
        for col in columns:
            self.result_tree.heading(col, text=col)
            if col == "Ticker":
                self.result_tree.column(col, width=50, anchor='w')
            elif col == "Current Price":
                self.result_tree.column(col, width=150, anchor='w')
            else:
                self.result_tree.column(col, width=200, anchor='w', stretch=True)

        self.result_tree.grid(row=0, column=0, sticky="nsew")

        self.result_tree.bind('<Configure>', lambda e: self.result_tree.column('#0', width=0, stretch=False))
        
        # Add scrollbars to the Treeview
        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        hsb = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self.setup_tags()

        self.root.update_idletasks()
        self.root.after(100, self.process_queue)

    def setup_tags(self):
        print("Setting up tags...")
        for color in ['green', 'red', 'black']:
            self.result_tree.tag_configure(color, foreground=self.tag_colors[color])
            print(f"Created tag: {color} with color {self.tag_colors[color]}")
        print("Tag setup complete.")

    def toggle_colorblind_mode(self):
        self.colorblind_mode = self.colorblind_var.get()
        self.setup_style()
        self.refresh_ui()

    def toggle_dark_mode(self):
        self.dark_mode = self.dark_mode_var.get()
        self.setup_style()
        self.setup_tags()
        self.refresh_ui()

    def refresh_ui(self):
        # Recursively update all widgets
        def update_widget(widget):
            widget_type = widget.winfo_class()
            if widget_type in ('TFrame', 'TLabel', 'TButton', 'TCheckbutton'):
                widget.configure(style=widget_type)
            for child in widget.winfo_children():
                update_widget(child)

        update_widget(self.root)

        # Refresh the Treeview
        self.result_tree.configure(style='Treeview')
        for item in self.result_tree.get_children():
            values = self.result_tree.item(item, 'values')
            tags = self.result_tree.item(item, 'tags')
            self.result_tree.delete(item)
            self.result_tree.insert('', 'end', values=values, tags=tags)

        # Reapply tags
        self.setup_tags()

    def calculate_row_height(self, values):
        font = ('Helvetica', 14)  # Match this with the font used in Treeview
        padding = 10  # Additional padding for each row

        max_lines = 1
        for value in values:
            lines = value.count('\n') + 1
            max_lines = max(max_lines, lines)

        line_height = font[1] + 2  # Font size plus a small buffer
        return (line_height * max_lines) + padding

    def validate_symbols(self, symbols):
        ticker_symbols = [symbol.strip().upper() for symbol in symbols.split(',') if symbol.strip()]
        invalid_symbols = [symbol for symbol in ticker_symbols if not symbol.isalnum()]
        if invalid_symbols:
            logging.error(f"Invalid ticker symbols entered: {', '.join(invalid_symbols)}")
            messagebox.showerror("Input Error", f"Invalid symbols: {', '.join(invalid_symbols)}")
            return []
        return ticker_symbols

    def run_analysis_method(self, method_function):
        symbols = self.validate_symbols(self.ticker_entry.get())
        if not symbols:
            return  # Exit the method if no valid symbols
        method_name = method_function.__name__
        column_index = {
            'calculate_value': 1,  # Buffett
            'evaluate_stock': 2,  # Brandes
            'calculate_hartz_millsap_hill': 3,  # Hartz
            'calculate_intrinsic_value': 4,  # Pabrai
            'calculate_hempton_nutty': 5
        }.get(method_name, 0)
        
        results = method_function(symbols, output=True)
        self.result_queue.put(('clear', None))
        for symbol in symbols:
            values = [''] * 7  # Increase to 7 columns to include price
            values[0] = symbol
            tags = ['black'] * 7  # Initialize all tags as black
            if symbol in results:
                result_text, tag = results[symbol]
                values[column_index] = result_text
                if column_index in [1, 3, 4]:  # Only for Buffett, Hartz, and Pabrai
                    tags[column_index] = tag
            
            # Always fetch current price and change
            current_price, price_change, price_change_percent = self.get_stock_price_and_change(symbol, self.period_var.get())
            if current_price is not None:
                price_info = f"Current: ${current_price:.2f}\nChange: ${price_change:.2f} ({price_change_percent:.2f}%)"
                values[6] = price_info  # Add price info to the last column
            else:
                values[6] = "Price data unavailable"
            
            self.result_queue.put(('insert', (values, tags)))

    def get_stock_price_and_change(self, symbol, period='1d'):
        try:
            stock = yf.Ticker(symbol)
            hist = stock.history(period=period)
            if hist.empty:
                return None, None, None
            current_price = hist['Close'].iloc[-1]
            open_price = hist['Open'].iloc[0]
            price_change = current_price - open_price
            price_change_percent = (price_change / open_price) * 100
            return current_price, price_change, price_change_percent
        except Exception as e:
            logging.error(f"Error fetching stock price for {symbol}: {e}")
            return None, None, None

    @lru_cache(maxsize=CACHE_SIZE)
    def fetch_stock_data(self, symbol):
        if symbol in self.data_cache:
            return self.data_cache[symbol]
        try:
            stock = yf.Ticker(symbol)
            data = {
                'financials': stock.financials,
                'balance_sheet': stock.balance_sheet,
                'info': stock.info,
                'cashflow': stock.cashflow.T
            }
            self.data_cache[symbol] = data
            logging.info(f"Fetched data for {symbol}")
            return data
        except Exception as e:
            logging.error(f"Error fetching data for {symbol}: {e}")
            messagebox.showerror("Data Fetch Error", f"An error occurred while fetching data for {symbol}: {e}")
            return None

    def clear_treeview(self):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

    def run_in_thread(self, func, *args):
        self.executor.submit(func, *args)

    def run_all_methods(self):
        symbols = self.validate_symbols(self.ticker_entry.get())
        if not symbols:
            return  # Exit the method if no valid symbols

        results = {symbol: {} for symbol in symbols}
        methods = [
            ("Buffett", self.calculate_value),
            ("Brandes", self.evaluate_stock),
            ("Hartz, Millsap, Hill", self.calculate_hartz_millsap_hill),
            ("Pabrai", self.calculate_intrinsic_value),
            ("Hempton Nutty", self.calculate_hempton_nutty)
        ]

        def run_method(method_name, method_function, symbols):
            method_results = method_function(symbols, output=True)
            for symbol, result in method_results.items():
                results[symbol][method_name] = result

        def run_methods():
            with ThreadPoolExecutor() as executor:
                futures = [executor.submit(run_method, method_name, method_function, symbols) for method_name, method_function in methods]
                for future in futures:
                    future.result()

            self.result_queue.put(('clear', None))
            for symbol in symbols:
                values = [symbol]
                tags = ['black'] * 7  # Always use black tags for Run All
                for method_name in ["Buffett", "Brandes", "Hartz, Millsap, Hill", "Pabrai", "Hempton Nutty"]:
                    result_tuple = results[symbol].get(method_name, ("No Data", 'black'))
                    result, _ = result_tuple  # Ignore the tag from individual methods
                    values.append(result)

                # Fetch current price and change
                current_price, price_change, price_change_percent = self.get_stock_price_and_change(symbol, self.period_var.get())
                if current_price is not None:
                    price_info = f"Current: ${current_price:.2f}\nChange: ${price_change:.2f} ({price_change_percent:.2f}%)"
                    values.append(price_info)
                else:
                    values.append("Price data unavailable")

                self.result_queue.put(('insert', (values, tags)))

        self.run_in_thread(run_methods)

    def process_queue(self):
        try:
            while True:
                action, data = self.result_queue.get_nowait()
                if action == 'clear':
                    self.clear_treeview()
                    print("Treeview cleared")
                elif action == 'insert':
                    values, tags = data
                    wrapped_values = [self.wrap_text(str(value), 30) for value in values]
                    
                    # Check if all tags are black (Run All case)
                    if all(tag == 'black' for tag in tags):
                        row_color = 'black'
                    else:
                        # Determine the row color based on Buffett, Hartz, and Pabrai methods
                        row_color = 'black'
                        color_indices = [1, 3, 4]  # Corresponding indices for Buffett, Hartz, and Pabrai
                        
                        for index in color_indices:
                            if tags[index] in ['green', 'red']:
                                row_color = tags[index]
                                break
                    
                    # Insert the item with the determined color tag
                    item = self.result_tree.insert('', 'end', values=wrapped_values, tags=(row_color,))
                    print(f"Inserted item: {item} with values: {wrapped_values} and color: {row_color}")

        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)

    def calculate_value(self, symbols, output):
        results = {}
        for symbol in symbols:
            data = self.fetch_stock_data(symbol)
            if not data:
                results[symbol] = (f"Error fetching data for {symbol}", 'red')
                continue
            try:
                income_statement = data['financials']
                balance_sheet = data['balance_sheet']
                pre_tax_income = income_statement.loc["Pretax Income"].head(5)
                average_pre_tax_income = pre_tax_income.mean()
                value_10x_pre_tax_income = average_pre_tax_income * 10
                adjusted_value = value_10x_pre_tax_income + balance_sheet.loc["Cash And Cash Equivalents"].iloc[0] - balance_sheet.loc["Total Debt"].iloc[0]
                market_cap = data['info']['marketCap']
                tag = 'green' if adjusted_value > market_cap else 'red'
                result_text = f"Adjusted Value:\n{adjusted_value:,.2f}\n\nMarket Cap:\n{market_cap:,.2f}"
                results[symbol] = (result_text, tag)
            except Exception as e:
                logging.error(f"Error calculating value for {symbol}: {e}")
                results[symbol] = (f"Error calculating value for {symbol}: {e}", 'red')
        if not output:
            self.result_queue.put(('clear', None))
            for symbol in symbols:
                result_text, tag = results[symbol]
                self.result_queue.put(('insert', ([symbol, result_text, '', '', '', ''], {symbol: tag})))
        return results

    def evaluate_stock(self, symbols, output):
        results = {}
        for symbol in symbols:
            data = self.fetch_stock_data(symbol)
            if not data:
                results[symbol] = (f"Error fetching data for {symbol}", 'red')
                continue
            try:
                financials = data['financials']
                balance_sheet = data['balance_sheet']
                info = data['info']
                net_income = financials.loc["Net Income"]
                no_losses = all(net_income.head(5) > 0)
                short_term_debt = balance_sheet.loc["Short Long Term Debt"].head(1).item() if 'Short Long Term Debt' in balance_sheet.index else 0
                long_term_debt = balance_sheet.loc["Long Term Debt"].head(1).item() if 'Long Term Debt' in balance_sheet.index else 0
                total_debt = short_term_debt + long_term_debt
                book_value = info['bookValue'] * info['sharesOutstanding']
                debt_to_equity = total_debt / book_value
                current_price = info['currentPrice']
                book_value_per_share = info['bookValue']
                price_to_book = current_price / book_value_per_share
                eps = info['trailingEps']
                earnings_yield = eps / current_price
                bond_yield = 0.09
                meets_bond_yield = earnings_yield > 2 * bond_yield

                result_text = f"No Losses: {'Yes' if no_losses else 'No'}\nDebt < 100% Equity: {'Yes' if debt_to_equity < 1 else 'No'}\nPrice < Book Value: {'Yes' if price_to_book < 1 else 'No'}\nEarnings Yield > 2x Bond Yield: {'Yes' if meets_bond_yield else 'No'}"
                
                # Determine overall tag
                if no_losses and debt_to_equity < 1 and price_to_book < 1 and meets_bond_yield:
                    tag = 'green'
                else:
                    tag = 'red'
                
                results[symbol] = (result_text, tag)
            except Exception as e:
                logging.error(f"Error evaluating stock for {symbol}: {e}")
                results[symbol] = (f"Error evaluating stock for {symbol}: {e}", 'red')
        if not output:
            self.result_queue.put(('clear', None))
            for symbol in symbols:
                result_text, tag = results[symbol]
                self.result_queue.put(('insert', ([symbol, '', result_text, '', '', ''], {symbol: tag})))
        return results

    def calculate_hartz_millsap_hill(self, symbols, output):
        results = {}
        for symbol in symbols:
            data = self.fetch_stock_data(symbol)
            if not data:
                results[symbol] = (f"Error fetching data for {symbol}", 'red')
                continue
            try:
                info = data['info']
                roe = info.get('returnOnEquity', 0)
                book_value_per_share = info.get('bookValue', 0)
                price_to_book_value = info.get('currentPrice', 0) / book_value_per_share if book_value_per_share else 0
                expected_returns = (roe / price_to_book_value) if price_to_book_value else 0
                tag = 'green' if expected_returns >= .07 else 'red'
                result_text = f"Expected Returns: {expected_returns * 100:.2f}%"
                results[symbol] = (result_text, tag)
            except Exception as e:
                logging.error(f"Error calculating Hartz, Millsap, Hill for {symbol}: {e}")
                results[symbol] = (f"Error calculating Hartz, Millsap, Hill for {symbol}: {e}", 'red')
        if not output:
            self.result_queue.put(('clear', None))
            for symbol in symbols:
                result_text, tag = results[symbol]
                self.result_queue.put(('insert', ([symbol, result_text, '', '', '', ''], {symbol: tag})))
        return results

    def calculate_intrinsic_value(self, symbols, output):
        results = {}
        for symbol in symbols:
            data = self.fetch_stock_data(symbol)
            if not data:
                results[symbol] = (f"Error fetching data for {symbol}", 'red')
                continue
            try:
                cash_flow = data['cashflow']
                balance_sheet = data['balance_sheet']
                free_cash_flow = cash_flow['Free Cash Flow']
                fcf_growth_rate = free_cash_flow.pct_change(fill_method=None).mean()
                future_fcf = free_cash_flow.iloc[-1] * (1 + fcf_growth_rate) ** 10
                cash_and_securities = balance_sheet.get('Cash And Cash Equivalents', [0])[0]
                other_short_term_investments = balance_sheet.get('Other Short Term Investments', [0])[0]
                intrinsic_value = future_fcf + cash_and_securities + other_short_term_investments
                market_cap = data['info']['marketCap']
                tag = 'green' if intrinsic_value > market_cap else ('red' if intrinsic_value < market_cap else 'black')
                result_text = f"Intrinsic Value: {intrinsic_value:,.2f}\nMarket Cap: {market_cap:,.2f}\nComparison: " + ("Undervalued" if intrinsic_value > market_cap else ("Overvalued" if intrinsic_value < market_cap else "Fairly Valued"))
                results[symbol] = (result_text, tag)
            except Exception as e:
                logging.error(f"Error calculating intrinsic value for {symbol}: {e}")
                results[symbol] = (f"Error calculating intrinsic value for {symbol}: {e}", 'red')
        if not output:
            self.result_queue.put(('clear', None))
            for symbol in symbols:
                result_text, tag = results[symbol]
                self.result_queue.put(('insert', ([symbol, '', '', '', result_text, ''], {symbol: tag})))
        return results

    def calculate_hempton_nutty(self, symbols, output):
        results = {}
        for symbol in symbols:
            data = self.fetch_stock_data(symbol)
            if not data:
                results[symbol] = (f"Error fetching data for {symbol}", 'red')
                continue
            try:
                info = data['info']
                market_cap = info.get('marketCap', 0)
                revenue = info.get('totalRevenue', 0) * 10
                ratio = market_cap / revenue if revenue else 0
                tag = 'green' if ratio <= 10 else 'red'
                result_text = f"Market Cap to 10x Revenue Ratio: {ratio:.2f}\nEvaluation: " + ("Bad Look (>10)" if ratio > 10 else "Acceptable (<=10)")
                results[symbol] = (result_text, tag)
            except Exception as e:
                logging.error(f"Error calculating Hempton Nutty for {symbol}: {e}")
                results[symbol] = (f"Error calculating Hempton Nutty for {symbol}: {e}", 'red')
        if not output:
            self.result_queue.put(('clear', None))
            for symbol in symbols:
                result_text, tag = results[symbol]
                self.result_queue.put(('insert', ([symbol, result_text, '', '', '', ''], {symbol: tag})))
        return results

if __name__ == "__main__":
    root = tk.Tk()
    app = StockAnalysisApp(root)
    root.mainloop()