# Designed for Windows and Linux Use
"""
Modbus RTU Device Discovery Tool
==============================================
GUI application for discovering Modbus RTU devices on various COM ports and baud rates.
This refactored version uses the ModbusMaster base class for clean, maintainable code.

Features:
- Multiple baud rate testing
- Device discovery and identification
- GUI with progress indication
- Detailed logging and results export

Author: Umair
Version: 1.0 
"""

import sys
import threading
import queue
import time
from typing import List, Tuple
from datetime import datetime
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import os

# Import our base Modbus class
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

try:
    from Templates.Modbus_Master_Base_v1 import ModbusMaster, SerialConfig, ModbusException
except ImportError:
    import tkinter.messagebox as mb
    mb.showerror("Missing Dependency", "modbus_master.py is required.\nPlease ensure it's in the same directory.")
    sys.exit(1)


class ModbusDiscoveryScanner(ModbusMaster):
    def __init__(self):
        super().__init__()
        self.stop_scan = False
        self.current_baudrate = None
    
    def configure_for_baudrate(self, port: str, baudrate: int, timeout: float = 0.2) -> bool:
        # Skip if already configured correctly
        if (self.current_baudrate == baudrate and 
            self.is_connected and 
            self.config and 
            self.config.port == port):
            return True
        
        try:
            # Disconnect if currently connected
            if self.is_connected:
                self.disconnect()
                time.sleep(0.1)
            
            # Create new configuration
            config = SerialConfig(
                port=port,
                baudrate=baudrate,
                timeout=timeout,
                write_timeout=timeout
            )
            
            # Connect with new configuration
            if self.connect(config):
                self.current_baudrate = baudrate
                return True
            
        except Exception as e:
            print(f"Configuration error: {e}")
        
        return False
    
    def scan_slaves_at_baudrate(self, port: str, baudrate: int, slave_range: range, 
                               timeout: float = 0.2, progress_callback=None) -> List[Tuple[int, str]]:
        found_devices = []
        
        # Configure for this baudrate
        if not self.configure_for_baudrate(port, baudrate, timeout):
            return found_devices
        
        # Scan slaves using base class method
        discovered = self.scan_slaves(slave_range, progress_callback)
        
        # Convert to expected format
        for slave_id, response in discovered:
            if not self.stop_scan:
                found_devices.append((slave_id, response.exception_name))
        
        return found_devices


class ModbusDiscoveryGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Modbus RTU Discovery Tool")
        self.root.geometry("900x750")
        self.root.minsize(800, 650)
        
        # Modern color scheme
        self.colors = {
            'bg': '#f8f9fa',
            'panel': '#ffffff',
            'accent': '#007bff',
            'success': '#28a745',
            'warning': '#ffc107',
            'danger': '#dc3545',
            'info': '#17a2b8',
            'dark': '#343a40',
            'light': '#e9ecef'
        }
        
        self.root.configure(bg=self.colors['bg'])
        
        # Set icon if available
        try:
            self.root.iconbitmap(default='icon.ico')
        except:
            pass
        
        # Initialize variables
        self.selected_port = tk.StringVar()
        self.slave_start = tk.IntVar(value=1)
        self.slave_end = tk.IntVar(value=247)
        self.timeout = tk.DoubleVar(value=0.2)
        self.scanning = False
        self.scanner = None
        self.scan_thread = None
        self.message_queue = queue.Queue()
        
        # Baud rate configuration
        self.baud_vars = {}
        self.default_bauds = [9600, 19200, 38400, 57600, 115200]
        
        # Create UI
        self.setup_styles()
        self.create_widgets()
        self.refresh_ports()
        self.process_queue()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_styles(self):
        """Configure modern ttk styles"""
        style = ttk.Style()
        
        # Use modern theme
        available_themes = style.theme_names()
        if 'vista' in available_themes:
            style.theme_use('vista')
        elif 'clam' in available_themes:
            style.theme_use('clam')
        
        # Configure custom styles
        style.configure("Title.TLabel", font=('Segoe UI', 14, 'bold'), foreground=self.colors['dark'])
        style.configure("Header.TLabel", font=('Segoe UI', 10, 'bold'), foreground=self.colors['dark'])
        style.configure("Info.TLabel", font=('Segoe UI', 8), foreground=self.colors['info'])
        style.configure("Success.TLabel", foreground=self.colors['success'], font=('Segoe UI', 9, 'bold'))
        style.configure("Error.TLabel", foreground=self.colors['danger'], font=('Segoe UI', 9, 'bold'))
        style.configure("Warning.TLabel", foreground=self.colors['warning'], font=('Segoe UI', 9, 'bold'))
    
    def create_widgets(self):
        """Create all GUI widgets with professional design"""
        # Main container with proper grid configuration
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)  # Make results section expandable
        
        # Create sections
        self.create_title_section(main_frame)
        self.create_config_section(main_frame)
        self.create_baudrate_section(main_frame)
        self.create_control_section(main_frame)
        self.create_results_section(main_frame)
        self.create_status_bar()
    
    def create_title_section(self, parent):
        """Create title section"""
        title_frame = tk.Frame(parent, bg=self.colors['bg'])
        title_frame.grid(row=0, column=0, sticky='ew', pady=(0, 20))
        
        ttk.Label(title_frame, text="🔍 Modbus RTU Device Discovery Tool", 
                 style="Title.TLabel").pack()
        ttk.Label(title_frame, text="Discover Modbus devices across multiple COM ports and baud rates", 
                 style="Info.TLabel").pack(pady=(5, 0))
    
    def create_config_section(self, parent):
        """Create configuration section"""
        config_frame = ttk.LabelFrame(parent, text="Connection Configuration", padding="15")
        config_frame.grid(row=1, column=0, sticky='ew', pady=(0, 15))
        config_frame.columnconfigure(1, weight=1)
        
        # COM Port selection row
        port_frame = tk.Frame(config_frame)
        port_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=5)
        port_frame.columnconfigure(1, weight=1)
        
        ttk.Label(port_frame, text="COM Port:", font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=(0, 10))
        
        port_input_frame = tk.Frame(port_frame)
        port_input_frame.grid(row=0, column=1, sticky='ew')
        port_input_frame.columnconfigure(0, weight=1)
        
        self.port_combo = ttk.Combobox(port_input_frame, textvariable=self.selected_port, 
                                      state='readonly', font=('Segoe UI', 9))
        self.port_combo.grid(row=0, column=0, sticky='ew', padx=(0, 10))
        
        ttk.Button(port_input_frame, text="🔄 Refresh", 
                  command=self.refresh_ports).grid(row=0, column=1)
        
        # Slave ID Range row
        slave_frame = tk.Frame(config_frame)
        slave_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=5)
        slave_frame.columnconfigure(1, weight=1)
        
        ttk.Label(slave_frame, text="Slave ID Range:", font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=(0, 10))
        
        range_input_frame = tk.Frame(slave_frame)
        range_input_frame.grid(row=0, column=1, sticky='e')
        
        ttk.Label(range_input_frame, text="From:").grid(row=0, column=0, padx=(0, 5))
        ttk.Spinbox(range_input_frame, from_=1, to=247, textvariable=self.slave_start, 
                   width=8, font=('Segoe UI', 9)).grid(row=0, column=1, padx=(0, 15))
        
        ttk.Label(range_input_frame, text="To:").grid(row=0, column=2, padx=(0, 5))
        ttk.Spinbox(range_input_frame, from_=1, to=247, textvariable=self.slave_end, 
                   width=8, font=('Segoe UI', 9)).grid(row=0, column=3)
        
        # Timeout row
        timeout_frame = tk.Frame(config_frame)
        timeout_frame.grid(row=2, column=0, columnspan=2, sticky='ew', pady=5)
        timeout_frame.columnconfigure(1, weight=1)
        
        ttk.Label(timeout_frame, text="Timeout (seconds):", font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=(0, 10))
        
        timeout_input_frame = tk.Frame(timeout_frame)
        timeout_input_frame.grid(row=0, column=1, sticky='e')
        
        ttk.Spinbox(timeout_input_frame, from_=0.1, to=5.0, increment=0.1, 
                   textvariable=self.timeout, width=8, 
                   font=('Segoe UI', 9)).grid(row=0, column=0, padx=(0, 5))
        
        ttk.Label(timeout_input_frame, text="(Increase for slow devices)", 
                 style="Info.TLabel").grid(row=0, column=1)
    
    def create_baudrate_section(self, parent):
        """Create baud rate selection section"""
        baud_frame = ttk.LabelFrame(parent, text="Baud Rates to Test", padding="15")
        baud_frame.grid(row=2, column=0, sticky='ew', pady=(0, 15))
        
        # Available baud rates
        baud_rates = [300, 600, 1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200, 230400]
        
        # Create checkboxes in rows
        for i, baud in enumerate(baud_rates):
            var = tk.BooleanVar(value=(baud in self.default_bauds))
            self.baud_vars[baud] = var
            
            row = i // 6
            col = i % 6
            
            cb = ttk.Checkbutton(baud_frame, text=f"{baud:,}", variable=var)
            cb.grid(row=row, column=col, padx=8, pady=3, sticky=tk.W)
        
        # Quick selection buttons
        button_frame = ttk.Frame(baud_frame)
        button_frame.grid(row=2, column=0, columnspan=6, pady=(15, 0))
        
        ttk.Button(button_frame, text="Select Common", 
                  command=self.select_common_bauds).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Select All", 
                  command=self.select_all_bauds).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear All", 
                  command=self.clear_all_bauds).pack(side=tk.LEFT, padx=5)
    
    def create_control_section(self, parent):
        """Create control buttons section"""
        control_frame = ttk.Frame(parent)
        control_frame.grid(row=3, column=0, pady=15)
        
        self.scan_btn = ttk.Button(control_frame, text="🔍 Start Discovery", 
                                  command=self.start_scan, style="Header.TLabel")
        self.scan_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(control_frame, text="⏹ Stop", 
                                  command=self.stop_scan, state="disabled")
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = ttk.Button(control_frame, text="🗑 Clear Results", 
                                   command=self.clear_results)
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = ttk.Button(control_frame, text="💾 Save Results", 
                                  command=self.save_results)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # Progress section within control area
        progress_frame = tk.Frame(control_frame, bg=self.colors['bg'])
        progress_frame.pack(side=tk.RIGHT, padx=(20, 0))
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate', length=200)
        self.progress.pack(side=tk.TOP)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to scan", 
                                       style="Info.TLabel")
        self.progress_label.pack(side=tk.TOP, pady=(2, 0))
    
    def create_results_section(self, parent):
        """Create results display section"""
        results_frame = ttk.LabelFrame(parent, text="Discovery Results", padding="10")
        results_frame.grid(row=4, column=0, sticky='nsew')
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Results text area with modern styling
        log_container = tk.Frame(results_frame, bg='#f8f9fa', relief='sunken', bd=1)
        log_container.grid(row=0, column=0, sticky='nsew', pady=(10, 0))
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)
        
        self.results_text = scrolledtext.ScrolledText(
            log_container,
            height=18,
            bg='#ffffff',
            fg='#333333',
            font=('Consolas', 9),
            relief='flat',
            bd=0,
            wrap='word'
        )
        self.results_text.grid(row=0, column=0, sticky='nsew', padx=2, pady=2)
        
        # Configure text tags for colored output
        self.results_text.tag_config("header", font=("Consolas", 9, "bold"), foreground="#0066cc")
        self.results_text.tag_config("success", foreground="#28a745", font=("Consolas", 9, "bold"))
        self.results_text.tag_config("error", foreground="#dc3545")
        self.results_text.tag_config("info", foreground="#007bff")
        self.results_text.tag_config("warning", foreground="#fd7e14")
    
    def create_status_bar(self):
        """Create status bar"""
        status_frame = tk.Frame(self.root, bg='#e9ecef', relief='sunken', bd=1)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_var = tk.StringVar(value="Ready - Select port and configure settings")
        self.status_label = tk.Label(status_frame, textvariable=self.status_var, 
                                   bg='#e9ecef', fg='#495057', anchor='w')
        self.status_label.pack(side=tk.LEFT, padx=10, pady=3)
        
        # Version info
        version_label = tk.Label(status_frame, text="v2.0", 
                                bg='#e9ecef', fg='#6c757d', anchor='e')
        version_label.pack(side=tk.RIGHT, padx=10, pady=3)
    
    # ==================== Port Management ====================
    
    def refresh_ports(self):
        """Refresh available COM ports using ModbusMaster"""
        try:
            ports = ModbusMaster.list_available_ports()
            port_list = []
            
            for port_info in ports:
                display_name = f"{port_info['port']} - {port_info['description']}"
                port_list.append(display_name)
            
            self.port_combo['values'] = port_list
            
            if port_list:
                self.port_combo.current(0)
                self.status_var.set(f"Found {len(port_list)} COM port(s)")
            else:
                self.status_var.set("No COM ports found")
                
            self.log_message(f"🔍 Found {len(ports)} COM port(s)", "info")
            
        except Exception as e:
            self.status_var.set(f"Error refreshing ports: {e}")
            self.log_message(f"❌ Error refreshing ports: {e}", "error")
    
    # ==================== Baud Rate Selection ====================
    
    def select_common_bauds(self):
        """Select common baud rates"""
        common = [9600, 19200, 38400]
        for baud, var in self.baud_vars.items():
            var.set(baud in common)
    
    def select_all_bauds(self):
        """Select all baud rates"""
        for var in self.baud_vars.values():
            var.set(True)
    
    def clear_all_bauds(self):
        """Clear all baud rates"""
        for var in self.baud_vars.values():
            var.set(False)
    
    def get_selected_bauds(self):
        """Get list of selected baud rates"""
        return [baud for baud, var in self.baud_vars.items() if var.get()]
    
    # ==================== Results Management ====================
    
    def clear_results(self):
        """Clear the results display"""
        self.results_text.delete(1.0, tk.END)
        self.progress['value'] = 0
        self.progress_label.config(text="Ready to scan")
        self.status_var.set("Results cleared")
    
    def log_message(self, message, tag=None):
        """Add message to results with optional formatting"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.results_text.insert(tk.END, formatted_message, tag)
        self.results_text.see(tk.END)
        self.root.update_idletasks()
    
    def save_results(self):
        """Save results to file"""
        if not hasattr(self, 'discovered_devices') or not self.discovered_devices:
            messagebox.showwarning("No Results", "No devices found to save")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"modbus_discovery_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write("MODBUS RTU DISCOVERY RESULTS\n")
                    f.write("=" * 50 + "\n\n")
                    f.write(f"Date/Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"Port: {self.selected_port.get().split(' - ')[0]}\n")
                    f.write(f"Slave Range: {self.slave_start.get()}-{self.slave_end.get()}\n")
                    f.write(f"Baud Rates: {self.get_selected_bauds()}\n")
                    f.write(f"Timeout: {self.timeout.get()}s\n")
                    f.write(f"Devices Found: {len(self.discovered_devices)}\n\n")
                    
                    f.write("DISCOVERED DEVICES:\n")
                    f.write("-" * 30 + "\n")
                    for device in self.discovered_devices:
                        f.write(f"Slave ID: {device['slave_id']}\n")
                        f.write(f"Baud Rate: {device['baud_rate']}\n")
                        f.write(f"Response: {device['exception_name']}\n")
                        f.write("\n")
                    
                    # Add full log
                    f.write("\n" + "="*50 + "\n")
                    f.write("FULL SCAN LOG:\n")
                    f.write("="*50 + "\n")
                    f.write(self.results_text.get(1.0, tk.END))
                
                messagebox.showinfo("Success", f"Results saved to:\n{filename}")
                self.status_var.set(f"Results saved to {os.path.basename(filename)}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save results:\n{str(e)}")
    
    # ==================== Scanning Operations ====================
    
    def start_scan(self):
        """Start the discovery scan"""
        if self.scanning:
            return
        
        # Validate inputs
        if not self.selected_port.get():
            messagebox.showerror("Error", "Please select a COM port")
            return
        
        selected_bauds = self.get_selected_bauds()
        if not selected_bauds:
            messagebox.showerror("Error", "Please select at least one baud rate")
            return
        
        if self.slave_start.get() > self.slave_end.get():
            messagebox.showerror("Error", "Invalid slave ID range")
            return
        
        # Clear previous results
        self.clear_results()
        
        # Update UI state
        self.scanning = True
        self.scan_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.refresh_btn.config(state="disabled")
        
        # Start scan thread
        port_str = self.selected_port.get().split(" - ")[0]
        slave_range = range(self.slave_start.get(), self.slave_end.get() + 1)
        
        self.scan_thread = threading.Thread(
            target=self.run_scan,
            args=(port_str, selected_bauds, slave_range),
            daemon=True
        )
        self.scan_thread.start()
    
    def stop_scan(self):
        """Stop the current scan"""
        if self.scanner:
            self.scanner.stop_scan = True
        self.status_var.set("Stopping scan...")
        self.log_message("⏹ Scan stop requested", "warning")
    
    def run_scan(self, port, baud_rates, slave_range):
        """Run the actual scan in separate thread"""
        try:
            # Initialize scanner
            self.scanner = ModbusDiscoveryScanner()
            all_devices = []
            
            # Log scan start
            self.message_queue.put(("header", "="*70))
            self.message_queue.put(("header", "🚀 STARTING MODBUS RTU DEVICE DISCOVERY"))
            self.message_queue.put(("header", "="*70))
            self.message_queue.put(("info", f"Port: {port}"))
            self.message_queue.put(("info", f"Baud rates: {baud_rates}"))
            self.message_queue.put(("info", f"Slave ID range: {slave_range.start}-{slave_range.stop-1}"))
            self.message_queue.put(("info", f"Timeout: {self.timeout.get()}s"))
            self.message_queue.put(("header", "="*70 + "\n"))
            
            total_operations = len(baud_rates) * len(slave_range)
            current_operation = 0
            
            for baud_rate in baud_rates:
                if self.scanner.stop_scan:
                    break
                
                self.message_queue.put(("info", f"🔍 Testing baud rate: {baud_rate:,}"))
                
                # Scan at this baud rate
                found_devices = self.scanner.scan_slaves_at_baudrate(
                    port, baud_rate, slave_range, self.timeout.get()
                )
                
                if found_devices:
                    self.message_queue.put(("success", f"   ✅ Found {len(found_devices)} device(s) at {baud_rate:,} baud"))
                    for slave_id, exception_name in found_devices:
                        self.message_queue.put(("success", f"      • Slave ID {slave_id}: {exception_name}"))
                        
                        device_info = {
                            'slave_id': slave_id,
                            'baud_rate': baud_rate,
                            'exception_name': exception_name
                        }
                        all_devices.append(device_info)
                else:
                    self.message_queue.put(("warning", f"   ⚠ No devices found at {baud_rate:,} baud"))
                
                # Update progress
                current_operation += len(slave_range)
                progress_pct = (current_operation / total_operations) * 100
                self.message_queue.put(("progress", progress_pct))
                
                self.message_queue.put(("", ""))  # Empty line
            
            # Generate summary
            self.message_queue.put(("header", "\n" + "="*70))
            self.message_queue.put(("header", "🎯 DISCOVERY COMPLETE"))
            self.message_queue.put(("header", "="*70))
            
            if all_devices:
                self.message_queue.put(("success", f"🎉 Found {len(all_devices)} Modbus device(s) total:"))
                for device in all_devices:
                    self.message_queue.put(("success", f"  • Slave ID {device['slave_id']} at {device['baud_rate']:,} baud"))
                
                self.discovered_devices = all_devices
            else:
                self.message_queue.put(("error", "❌ No Modbus devices found"))
                self.message_queue.put(("info", "\n💡 Troubleshooting tips:"))
                self.message_queue.put(("info", "  • Check device connections and power"))
                self.message_queue.put(("info", "  • Verify RS485 wiring (A/B terminals)"))
                self.message_queue.put(("info", "  • Try different baud rates"))
                self.message_queue.put(("info", "  • Increase timeout for slow devices"))
                self.message_queue.put(("info", "  • Check if port is used by another program"))
                self.discovered_devices = []
            
        except Exception as e:
            self.message_queue.put(("error", f"❌ Error during scan: {str(e)}"))
        finally:
            if self.scanner:
                self.scanner.disconnect()
            self.message_queue.put(("done", None))
    
    def process_queue(self):
        """Process messages from scan thread"""
        try:
            while True:
                tag, message = self.message_queue.get_nowait()
                
                if tag == "progress":
                    self.progress['value'] = message
                elif tag == "status":
                    self.progress_label.config(text=message)
                elif tag == "done":
                    # Scan completed
                    self.scanning = False
                    self.scan_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    self.refresh_btn.config(state="normal")
                    self.progress['value'] = 100
                    self.progress_label.config(text="Scan complete")
                    self.status_var.set("Discovery complete")
                    if self.scanner:
                        self.scanner.stop_scan = False
                else:
                    if message:  # Don't log empty messages
                        self.log_message(message, tag if tag else None)
                        
        except queue.Empty:
            pass
        finally:
            # Schedule next queue check
            self.root.after(100, self.process_queue)
    
    def on_closing(self):
        """Handle window close event"""
        if self.scanning:
            if messagebox.askokcancel("Quit", "A scan is in progress. Stop it and exit?"):
                self.stop_scan()
                time.sleep(0.5)
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    """Main entry point"""
    try:
        root = tk.Tk()
        app = ModbusDiscoveryGUI(root)
        
        # Center window on screen
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        root.mainloop()
        
    except Exception as e:
        print(f"Application error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()