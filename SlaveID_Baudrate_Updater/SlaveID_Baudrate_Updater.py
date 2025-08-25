"""
Modbus RTU Device Configuration Tool - Refactored
=================================================

A professional tool for configuring Modbus RTU device parameters using
block read-modify-write operations for safe configuration changes.
This refactored version uses the ModbusMaster base class.

Features:
- Safe block read-modify-write operations
- Device communication testing
- Modern GUI with real-time feedback
- Comprehensive error handling
- Cross-platform compatibility

Author: Umair
Version: 1.0 
"""

import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import time
from datetime import datetime
from typing import Optional, Tuple
import os

# Import base Modbus class
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

try:
    from Templates.Modbus_Master_Base_v1 import ModbusMaster, SerialConfig, ModbusException
except ImportError:
    import tkinter.messagebox as mb
    mb.showerror("Missing Dependency", "modbus_master.py is required.\nPlease ensure it's in the same directory.")
    import sys
    sys.exit(1)


class ModbusConfigurationTool(ModbusMaster):
    def __init__(self):
        super().__init__()
        self.connection_tested = False
        
        # Common register mappings for different device types
        self.device_profiles = {
            'Generic': {
                'slave_id_register': 2048,
                'baudrate_register': 2051,
                'description': 'Generic Modbus device'
            },
            'Custom Device 1': {
                'slave_id_register': 1000,
                'baudrate_register': 1001,
                'description': 'Custom device type 1'
            },
            'Custom Device 2': {
                'slave_id_register': 40001,
                'baudrate_register': 40002,
                'description': 'Custom device type 2'
            }
        }
    
    def test_comprehensive_communication(self, slave_id: int) -> Tuple[bool, str, dict]:
        """
        Perform comprehensive communication test with detailed results
        
        Args:
            slave_id: Slave ID to test
            
        Returns:
            Tuple of (success, message, test_results_dict)
        """
        if not self.is_connected:
            return False, "Not connected", {}
        
        test_results = {
            'holding_registers': [],
            'input_registers': [],
            'successful_addresses': [],
            'failed_addresses': [],
            'response_times': []
        }
        
        # Test addresses to try
        test_addresses = [0, 1, 100, 1000, 2048, 2051, 40001, 30001]
        
        successful_tests = 0
        
        for addr in test_addresses:
            start_time = time.time()
            
            # Test holding registers
            try:
                registers, error = self.read_holding_registers(slave_id, addr, 1)
                response_time = (time.time() - start_time) * 1000  # ms
                
                if registers is not None:
                    test_results['holding_registers'].append({
                        'address': addr,
                        'value': registers[0],
                        'response_time': response_time
                    })
                    test_results['successful_addresses'].append(addr)
                    successful_tests += 1
                else:
                    test_results['failed_addresses'].append(addr)
                    
            except Exception:
                test_results['failed_addresses'].append(addr)
            
            # Test input registers if holding registers failed
            if addr in test_results['failed_addresses']:
                try:
                    registers, error = self.read_input_registers(slave_id, addr, 1)
                    if registers is not None:
                        test_results['input_registers'].append({
                            'address': addr,
                            'value': registers[0],
                            'response_time': response_time
                        })
                        test_results['successful_addresses'].append(addr)
                        test_results['failed_addresses'].remove(addr)
                        successful_tests += 1
                except:
                    pass
        
        if successful_tests > 0:
            avg_response_time = sum([r['response_time'] for r in test_results['holding_registers'] + test_results['input_registers']]) / successful_tests
            message = f"Communication successful! {successful_tests} addresses responded (avg: {avg_response_time:.1f}ms)"
            self.connection_tested = True
            return True, message, test_results
        else:
            message = f"No response from slave ID {slave_id}. Check connections and settings."
            return False, message, test_results
    
    def block_read_modify_write(self, slave_id: int, register_updates: dict) -> Tuple[bool, str]:
        """
        Perform block read-modify-write operation for multiple registers
        
        Args:
            slave_id: Target slave ID
            register_updates: Dictionary of {register_address: new_value}
            
        Returns:
            Tuple of (success, error_message)
        """
        if not register_updates:
            return False, "No register updates specified"
        
        try:
            # Determine the address range to read/write
            addresses = list(register_updates.keys())
            start_address = min(addresses)
            end_address = max(addresses)
            register_count = (end_address - start_address) + 1
            
            # Read current values
            current_values, error = self.read_holding_registers(slave_id, start_address, register_count)
            
            if current_values is None:
                return False, f"Failed to read current values: {error}"
            
            # Create a copy for modification
            modified_values = current_values.copy()
            
            # Apply updates to specific registers
            changes_made = []
            for reg_addr, new_value in register_updates.items():
                index = reg_addr - start_address
                if 0 <= index < len(modified_values):
                    old_value = modified_values[index]
                    modified_values[index] = new_value
                    changes_made.append(f"Register {reg_addr}: {old_value} → {new_value}")
            
            # Write the entire block back
            success, error = self.write_multiple_registers(slave_id, start_address, modified_values)
            
            if success:
                return True, f"Block write successful. Changes: {'; '.join(changes_made)}"
            else:
                return False, f"Block write failed: {error}"
                
        except Exception as e:
            return False, f"Block operation error: {str(e)}"
    
    def validate_configuration(self, slave_id: int, new_slave_id: int, new_baudrate: int) -> Tuple[bool, str]:
        """
        Validate configuration parameters before writing
        
        Args:
            slave_id: Current slave ID
            new_slave_id: New slave ID to set
            new_baudrate: New baud rate to set
            
        Returns:
            Tuple of (valid, error_message)
        """
        # Validate slave ID range
        if not (1 <= new_slave_id <= 247):
            return False, "Slave ID must be between 1 and 247"
        
        # Validate baud rate
        if not (300 <= new_baudrate <= 1000000):
            return False, "Baud rate must be between 300 and 1,000,000"
        
        # Check if new slave ID conflicts with existing devices (if possible)
        if new_slave_id != slave_id:
            # Quick test to see if new slave ID is already in use
            test_response = self.probe_device(new_slave_id)
            if test_response.is_valid:
                return False, f"Slave ID {new_slave_id} appears to already be in use"
        
        return True, "Configuration valid"


class ModbusConfigGUI:
    """
    Modern GUI for Modbus device configuration with enhanced user experience
    """
    
    def __init__(self, root):
        self.root = root
        self.root.title("Modbus RTU Configuration Tool")
        self.root.geometry("800x700")
        self.root.minsize(700, 600)
        
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
        
        # Initialize components
        self.config_tool: Optional[ModbusConfigurationTool] = None
        self.is_connected = False
        
        # GUI variables
        self.port_var = tk.StringVar()
        self.baudrate_var = tk.StringVar(value="9600")
        self.current_slave_var = tk.StringVar(value="1")
        self.slave_id_reg_var = tk.StringVar(value="2048")
        self.new_slave_id_var = tk.StringVar(value="1")
        self.baudrate_reg_var = tk.StringVar(value="2051")
        self.new_baudrate_var = tk.StringVar(value="9600")
        self.device_profile_var = tk.StringVar(value="Generic")
        
        # Create UI
        self.setup_styles()
        self.create_widgets()
        self.refresh_ports()
        
        # Handle window closing
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
        """Create all GUI widgets"""
        # Main container with proper grid configuration
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)  # Make log section expandable
        
        # Create sections
        self.create_title_section(main_frame)
        self.create_connection_section(main_frame)
        self.create_configuration_section(main_frame)
        self.create_log_section(main_frame)
        self.create_status_bar()
    
    def create_title_section(self, parent):
        """Create title section"""
        title_frame = tk.Frame(parent, bg=self.colors['bg'])
        title_frame.grid(row=0, column=0, sticky='ew', pady=(0, 20))
        
        ttk.Label(title_frame, text="⚙️ Modbus RTU Configuration Tool", 
                 style="Title.TLabel").pack()
        ttk.Label(title_frame, text="Safe device configuration using block read-modify-write operations", 
                 style="Info.TLabel").pack(pady=(5, 0))
    
    def create_connection_section(self, parent):
        """Create connection configuration section"""
        conn_frame = ttk.LabelFrame(parent, text="Connection Settings", padding="15")
        conn_frame.grid(row=1, column=0, sticky='ew', pady=(0, 15))
        conn_frame.columnconfigure(1, weight=1)
        
        # Port selection row
        port_frame = tk.Frame(conn_frame)
        port_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=5)
        port_frame.columnconfigure(1, weight=1)
        
        ttk.Label(port_frame, text="COM Port:", font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=(0, 10))
        
        port_input_frame = tk.Frame(port_frame)
        port_input_frame.grid(row=0, column=1, sticky='ew')
        port_input_frame.columnconfigure(0, weight=1)
        
        self.port_combo = ttk.Combobox(port_input_frame, textvariable=self.port_var, 
                                      state='readonly', font=('Segoe UI', 9))
        self.port_combo.grid(row=0, column=0, sticky='ew', padx=(0, 10))
        
        ttk.Button(port_input_frame, text="🔄 Refresh", 
                  command=self.refresh_ports).grid(row=0, column=1)
        
        # Baudrate row
        baud_frame = tk.Frame(conn_frame)
        baud_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=5)
        baud_frame.columnconfigure(1, weight=1)
        
        ttk.Label(baud_frame, text="Baudrate:", font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=(0, 10))
        
        baud_combo = ttk.Combobox(baud_frame, textvariable=self.baudrate_var,
                                 values=["1200", "2400", "4800", "9600", "19200", "38400", "57600", "115200"],
                                 state='readonly', font=('Segoe UI', 9))
        baud_combo.grid(row=0, column=1, sticky='e')
        
        # Current slave ID row
        slave_frame = tk.Frame(conn_frame)
        slave_frame.grid(row=2, column=0, columnspan=2, sticky='ew', pady=5)
        slave_frame.columnconfigure(1, weight=1)
        
        ttk.Label(slave_frame, text="Current Slave ID:", font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=(0, 10))
        ttk.Entry(slave_frame, textvariable=self.current_slave_var, 
                 font=('Segoe UI', 9), width=10).grid(row=0, column=1, sticky='e')
        
        # Connection controls
        control_frame = tk.Frame(conn_frame)
        control_frame.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(15, 5))
        
        self.connect_btn = ttk.Button(control_frame, text="🔌 Connect", 
                                     command=self.toggle_connection, style="Header.TLabel")
        self.connect_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.test_btn = ttk.Button(control_frame, text="🧪 Test Communication", 
                                  command=self.test_communication, state="disabled")
        self.test_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Connection status
        self.status_var = tk.StringVar(value="⚫ Disconnected")
        self.status_label = ttk.Label(control_frame, textvariable=self.status_var, 
                                     style="Error.TLabel")
        self.status_label.pack(side=tk.RIGHT)
    
    def create_configuration_section(self, parent):
        """Create device configuration section"""
        config_frame = ttk.LabelFrame(parent, text="Device Configuration", padding="15")
        config_frame.grid(row=2, column=0, sticky='ew', pady=(0, 15))
        config_frame.columnconfigure(1, weight=1)
        
        # Device profile selection
        profile_frame = tk.Frame(config_frame)
        profile_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        profile_frame.columnconfigure(1, weight=1)
        
        ttk.Label(profile_frame, text="Device Profile:", font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=(0, 10))
        
        self.profile_combo = ttk.Combobox(profile_frame, textvariable=self.device_profile_var,
                                         values=list(ModbusConfigurationTool().device_profiles.keys()),
                                         state='readonly', font=('Segoe UI', 9))
        self.profile_combo.grid(row=0, column=1, sticky='ew')
        self.profile_combo.bind('<<ComboboxSelected>>', self.on_profile_selected)
        
        # Register settings grid
        grid_frame = tk.Frame(config_frame)
        grid_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(10, 0))
        grid_frame.columnconfigure(1, weight=1)
        
        # Slave ID register
        self.create_config_row(grid_frame, "Slave ID Register:", self.slave_id_reg_var, 0)
        self.create_config_row(grid_frame, "New Slave ID:", self.new_slave_id_var, 1)
        
        # Separator
        ttk.Separator(grid_frame, orient='horizontal').grid(row=2, column=0, columnspan=2, 
                                                           sticky='ew', pady=10)
        
        # Baudrate register
        self.create_config_row(grid_frame, "Baudrate Register:", self.baudrate_reg_var, 3)
        
        # New baudrate with validation
        baud_row_frame = tk.Frame(grid_frame)
        baud_row_frame.grid(row=4, column=0, columnspan=2, sticky='ew', pady=8)
        baud_row_frame.columnconfigure(1, weight=1)
        
        ttk.Label(baud_row_frame, text="New Baudrate:", font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=(0, 10))
        
        baud_input_frame = tk.Frame(baud_row_frame)
        baud_input_frame.grid(row=0, column=1, sticky='e')
        
        self.new_baud_combo = ttk.Combobox(baud_input_frame, textvariable=self.new_baudrate_var,
                                          values=["1200", "2400", "4800", "9600", "19200", "38400", "57600", "115200"],
                                          font=('Segoe UI', 9), width=12)
        self.new_baud_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(baud_input_frame, text="(editable)", style="Info.TLabel").pack(side=tk.LEFT)
        
        # Action buttons
        action_frame = tk.Frame(config_frame)
        action_frame.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(20, 0))
        
        self.read_btn = ttk.Button(action_frame, text="📖 Read Current Values", 
                                  command=self.read_current_values, state="disabled")
        self.read_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.validate_btn = ttk.Button(action_frame, text="✅ Validate Settings", 
                                      command=self.validate_settings, state="disabled")
        self.validate_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.write_btn = ttk.Button(action_frame, text="💾 Write Configuration", 
                                   command=self.write_configuration, state="disabled",
                                   style="Header.TLabel")
        self.write_btn.pack(side=tk.LEFT)
    
    def create_config_row(self, parent, label_text, var, row):
        """Create a configuration input row"""
        ttk.Label(parent, text=label_text, font=('Segoe UI', 9)).grid(row=row, column=0, sticky='w', padx=(0, 10), pady=5)
        ttk.Entry(parent, textvariable=var, font=('Segoe UI', 9), width=12).grid(row=row, column=1, sticky='e', pady=5)
    
    def create_log_section(self, parent):
        """Create activity log section"""
        log_frame = ttk.LabelFrame(parent, text="Activity Log", padding="10")
        log_frame.grid(row=3, column=0, sticky='nsew')
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Log text with modern styling
        log_container = tk.Frame(log_frame, bg='#f8f9fa', relief='sunken', bd=1)
        log_container.grid(row=0, column=0, sticky='nsew', pady=(10, 0))
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(
            log_container,
            height=12,
            bg='#ffffff',
            fg='#333333',
            font=('Consolas', 9),
            relief='flat',
            bd=0,
            wrap='word'
        )
        self.log_text.grid(row=0, column=0, sticky='nsew', padx=2, pady=2)
        
        # Configure text tags for colored logging
        self.log_text.tag_config("info", foreground=self.colors['info'])
        self.log_text.tag_config("success", foreground=self.colors['success'], font=('Consolas', 9, 'bold'))
        self.log_text.tag_config("warning", foreground=self.colors['warning'])
        self.log_text.tag_config("error", foreground=self.colors['danger'], font=('Consolas', 9, 'bold'))
        self.log_text.tag_config("header", foreground=self.colors['dark'], font=('Consolas', 9, 'bold'))
        
        # Log controls
        log_controls = tk.Frame(log_frame)
        log_controls.grid(row=1, column=0, sticky='ew', pady=(10, 0))
        
        ttk.Button(log_controls, text="🗑️ Clear Log", 
                  command=self.clear_log).pack(side=tk.RIGHT)
        
        ttk.Button(log_controls, text="💾 Save Log", 
                  command=self.save_log).pack(side=tk.RIGHT, padx=(0, 10))
    
    def create_status_bar(self):
        """Create status bar"""
        status_frame = tk.Frame(self.root, bg='#e9ecef', relief='sunken', bd=1)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_bar_label = tk.Label(status_frame, text="Ready - Configure connection and device settings", 
                                        bg='#e9ecef', fg='#495057', anchor='w')
        self.status_bar_label.pack(side=tk.LEFT, padx=10, pady=3)
        
        # Version info
        version_label = tk.Label(status_frame, text="v2.0", 
                                bg='#e9ecef', fg='#6c757d', anchor='e')
        version_label.pack(side=tk.RIGHT, padx=10, pady=3)
    
    # ==================== Event Handlers ====================
    
    def on_profile_selected(self, event=None):
        """Handle device profile selection"""
        profile_name = self.device_profile_var.get()
        if profile_name and self.config_tool:
            profile = self.config_tool.device_profiles.get(profile_name, {})
            
            # Update register addresses
            if 'slave_id_register' in profile:
                self.slave_id_reg_var.set(str(profile['slave_id_register']))
            if 'baudrate_register' in profile:
                self.baudrate_reg_var.set(str(profile['baudrate_register']))
            
            self.log(f"📋 Applied device profile: {profile_name}")
            if 'description' in profile:
                self.log(f"   Description: {profile['description']}", "info")
    
    # ==================== Port Management ====================
    
    def refresh_ports(self):
        """Refresh available COM ports"""
        try:
            ports = ModbusMaster.list_available_ports()
            port_list = []
            
            for port_info in ports:
                display_name = f"{port_info['port']} - {port_info['description']}"
                port_list.append(display_name)
            
            self.port_combo['values'] = port_list
            
            if port_list:
                self.port_combo.current(0)
                self.status_bar_label.config(text=f"Found {len(port_list)} COM port(s)")
            else:
                self.status_bar_label.config(text="No COM ports found")
                
            self.log(f"🔍 Found {len(ports)} COM port(s)")
            
        except Exception as e:
            self.status_bar_label.config(text=f"Error refreshing ports: {e}")
            self.log(f"❌ Error refreshing ports: {e}", "error")
    
    # ==================== Connection Management ====================
    
    def toggle_connection(self):
        """Toggle connection state"""
        if self.is_connected:
            self.disconnect()
        else:
            self.connect()
    
    def connect(self):
        """Connect to Modbus device"""
        port_selection = self.port_var.get()
        if not port_selection:
            messagebox.showerror("Error", "Please select a COM port")
            return
        
        port = port_selection.split(" - ")[0]  # Extract port name
        baudrate_str = self.baudrate_var.get()
        
        if not baudrate_str:
            messagebox.showerror("Error", "Please select a baudrate")
            return
        
        try:
            baudrate = int(baudrate_str)
            
            # Disable connect button during connection
            self.connect_btn.config(state="disabled", text="🔄 Connecting...")
            self.log(f"🔌 Attempting to connect to {port} at {baudrate:,} baud...")
            
            # Run connection in separate thread
            connection_thread = threading.Thread(
                target=self.connect_worker, 
                args=(port, baudrate), 
                daemon=True
            )
            connection_thread.start()
            
        except ValueError:
            messagebox.showerror("Error", "Invalid baudrate")
            self.connect_btn.config(state="normal", text="🔌 Connect")
    
    def connect_worker(self, port, baudrate):
        """Worker thread for connection"""
        try:
            # Create configuration tool
            self.config_tool = ModbusConfigurationTool()
            
            # Create serial configuration
            config = SerialConfig(port=port, baudrate=baudrate, timeout=2.0)
            
            # Connect
            if self.config_tool.connect(config):
                self.root.after(0, self.connection_success, port, baudrate)
            else:
                self.root.after(0, self.connection_failed, "Failed to open serial connection")
                
        except Exception as e:
            self.root.after(0, self.connection_failed, str(e))
    
    def connection_success(self, port, baudrate):
        """Handle successful connection"""
        self.is_connected = True
        self.connect_btn.config(state="normal", text="🔌 Disconnect")
        self.test_btn.config(state="normal")
        self.read_btn.config(state="normal")
        self.validate_btn.config(state="normal")
        self.status_var.set("🟢 Connected")
        self.status_label.config(style="Success.TLabel")
        self.status_bar_label.config(text=f"Connected to {port} at {baudrate:,} baud")
        
        self.log(f"✅ Successfully connected to {port} at {baudrate:,} baud", "success")
    
    def connection_failed(self, error_msg):
        """Handle connection failure"""
        self.is_connected = False
        self.connect_btn.config(state="normal", text="🔌 Connect")
        self.test_btn.config(state="disabled")
        self.read_btn.config(state="disabled")
        self.validate_btn.config(state="disabled")
        self.write_btn.config(state="disabled")
        self.status_var.set("⚫ Disconnected")
        self.status_label.config(style="Error.TLabel")
        self.status_bar_label.config(text="Connection failed")
        
        self.log(f"❌ Connection failed: {error_msg}", "error")
        
        # Clean up
        if self.config_tool:
            try:
                self.config_tool.disconnect()
            except:
                pass
            self.config_tool = None
    
    def disconnect(self):
        """Disconnect from device"""
        try:
            if self.config_tool:
                self.config_tool.disconnect()
                self.config_tool = None
            
            self.is_connected = False
            self.connect_btn.config(text="🔌 Connect")
            self.test_btn.config(state="disabled")
            self.read_btn.config(state="disabled")
            self.validate_btn.config(state="disabled")
            self.write_btn.config(state="disabled")
            self.status_var.set("⚫ Disconnected")
            self.status_label.config(style="Error.TLabel")
            self.status_bar_label.config(text="Disconnected")
            
            self.log("🔌 Disconnected from device", "info")
            
        except Exception as e:
            self.log(f"❌ Disconnect error: {e}", "error")
    
    # ==================== Communication Testing ====================
    
    def test_communication(self):
        """Test communication with device"""
        if not self.is_connected or not self.config_tool:
            messagebox.showerror("Error", "Not connected to device")
            return
        
        try:
            current_slave_id = int(self.current_slave_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid slave ID")
            return
        
        self.test_btn.config(state="disabled", text="🔄 Testing...")
        self.log(f"🧪 Testing communication with Slave ID {current_slave_id}...")
        
        # Run test in separate thread
        test_thread = threading.Thread(
            target=self.test_communication_worker,
            args=(current_slave_id,),
            daemon=True
        )
        test_thread.start()
    
    def test_communication_worker(self, slave_id):
        """Worker thread for communication testing"""
        try:
            success, message, test_results = self.config_tool.test_comprehensive_communication(slave_id)
            self.root.after(0, self.test_communication_complete, success, message, test_results)
        except Exception as e:
            self.root.after(0, self.test_communication_complete, False, str(e), {})
    
    def test_communication_complete(self, success, message, test_results):
        """Handle test completion"""
        self.test_btn.config(state="normal", text="🧪 Test Communication")
        
        if success:
            self.log(f"✅ {message}", "success")
            self.write_btn.config(state="normal")
            
            # Log detailed results
            if test_results.get('holding_registers'):
                self.log("   📊 Accessible holding registers:", "info")
                for reg in test_results['holding_registers']:
                    self.log(f"      Register {reg['address']}: {reg['value']} (response: {reg['response_time']:.1f}ms)")
            
            if test_results.get('input_registers'):
                self.log("   📊 Accessible input registers:", "info")
                for reg in test_results['input_registers']:
                    self.log(f"      Register {reg['address']}: {reg['value']} (response: {reg['response_time']:.1f}ms)")
        else:
            self.log(f"❌ {message}", "error")
            self.write_btn.config(state="disabled")
            
            # Suggestions for troubleshooting
            self.log("   💡 Troubleshooting suggestions:", "info")
            self.log("      • Verify slave ID is correct")
            self.log("      • Check RS485 wiring (A/B terminals)")
            self.log("      • Try different baud rate")
            self.log("      • Ensure device is powered and functioning")
    
    # ==================== Configuration Operations ====================
    
    def read_current_values(self):
        """Read current register values"""
        if not self.is_connected or not self.config_tool:
            messagebox.showerror("Error", "Not connected to device")
            return
        
        try:
            current_slave_id = int(self.current_slave_var.get())
            slave_reg = int(self.slave_id_reg_var.get())
            baud_reg = int(self.baudrate_reg_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid register addresses or slave ID")
            return
        
        self.read_btn.config(state="disabled", text="📖 Reading...")
        self.log(f"📖 Reading current values from registers...")
        
        # Run read in separate thread
        read_thread = threading.Thread(
            target=self.read_current_values_worker,
            args=(current_slave_id, slave_reg, baud_reg),
            daemon=True
        )
        read_thread.start()
    
    def read_current_values_worker(self, slave_id, slave_reg, baud_reg):
        """Worker thread for reading values"""
        try:
            # Determine range to read
            start_reg = min(slave_reg, baud_reg)
            end_reg = max(slave_reg, baud_reg)
            count = (end_reg - start_reg) + 1
            
            # Read registers
            registers, error = self.config_tool.read_holding_registers(slave_id, start_reg, count)
            
            self.root.after(0, self.read_complete, registers, error, start_reg, slave_reg, baud_reg)
            
        except Exception as e:
            self.root.after(0, self.read_complete, None, str(e), 0, 0, 0)
    
    def read_complete(self, registers, error, start_reg, slave_reg, baud_reg):
        """Handle read completion"""
        self.read_btn.config(state="normal", text="📖 Read Current Values")
        
        if registers is not None:
            self.log(f"✅ Successfully read {len(registers)} registers starting from {start_reg}", "success")
            
            # Display specific register values
            for i, value in enumerate(registers):
                reg_addr = start_reg + i
                if reg_addr == slave_reg:
                    self.log(f"   🆔 Slave ID Register ({reg_addr}): {value}")
                elif reg_addr == baud_reg:
                    self.log(f"   ⚡ Baudrate Register ({reg_addr}): {value}")
                else:
                    self.log(f"   📊 Register {reg_addr}: {value}")
        else:
            self.log(f"❌ Failed to read registers: {error}", "error")
    
    def validate_settings(self):
        """Validate configuration settings"""
        if not self.is_connected or not self.config_tool:
            messagebox.showerror("Error", "Not connected to device")
            return
        
        try:
            current_slave_id = int(self.current_slave_var.get())
            new_slave_id = int(self.new_slave_id_var.get())
            new_baudrate = int(self.new_baudrate_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid configuration values")
            return
        
        self.log("✅ Validating configuration settings...")
        
        # Validate using the configuration tool
        valid, message = self.config_tool.validate_configuration(current_slave_id, new_slave_id, new_baudrate)
        
        if valid:
            self.log(f"✅ Configuration valid: {message}", "success")
            messagebox.showinfo("Validation", "✅ Configuration settings are valid!")
        else:
            self.log(f"❌ Configuration invalid: {message}", "error")
            messagebox.showerror("Validation Error", f"❌ {message}")
    
    def write_configuration(self):
        """Write configuration to device"""
        if not self.is_connected or not self.config_tool:
            messagebox.showerror("Error", "Not connected to device")
            return
        
        if not self.config_tool.connection_tested:
            if not messagebox.askyesno("Warning", 
                                     "Communication hasn't been tested. Continue anyway?"):
                return
        
        try:
            current_slave_id = int(self.current_slave_var.get())
            new_slave_id = int(self.new_slave_id_var.get())
            slave_reg = int(self.slave_id_reg_var.get())
            new_baudrate = int(self.new_baudrate_var.get())
            baud_reg = int(self.baudrate_reg_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid configuration values")
            return
        
        # Final confirmation
        confirm_msg = (f"⚠️ CONFIRMATION REQUIRED ⚠️\n\n"
                      f"This will modify device configuration:\n"
                      f"• Current Slave ID: {current_slave_id} → New: {new_slave_id}\n"
                      f"• New Baudrate: {new_baudrate:,}\n\n"
                      f"Register updates:\n"
                      f"• Register {slave_reg}: {new_slave_id}\n"
                      f"• Register {baud_reg}: {new_baudrate}\n\n"
                      f"Device may need power cycle after configuration.\n\n"
                      f"Continue with configuration write?")
        
        if not messagebox.askyesno("Confirm Configuration Write", confirm_msg):
            return
        
        self.write_btn.config(state="disabled", text="💾 Writing...")
        self.log("🚀 Starting configuration write using block read-modify-write...")
        
        # Prepare register updates
        register_updates = {
            slave_reg: new_slave_id,
            baud_reg: new_baudrate
        }
        
        # Run write in separate thread
        write_thread = threading.Thread(
            target=self.write_configuration_worker,
            args=(current_slave_id, register_updates),
            daemon=True
        )
        write_thread.start()
    
    def write_configuration_worker(self, slave_id, register_updates):
        """Worker thread for configuration writing"""
        try:
            success, message = self.config_tool.block_read_modify_write(slave_id, register_updates)
            self.root.after(0, self.write_complete, success, message)
        except Exception as e:
            self.root.after(0, self.write_complete, False, str(e))
    
    def write_complete(self, success, message):
        """Handle write completion"""
        self.write_btn.config(state="normal", text="💾 Write Configuration")
        
        if success:
            self.log(f"🎉 Configuration write successful!", "success")
            self.log(f"   Details: {message}", "info")
            
            messagebox.showinfo("Success", 
                              "✅ Configuration written successfully!\n\n"
                              "📋 Next steps:\n"
                              "• Power cycle the device for changes to take effect\n"
                              "• Update connection settings if slave ID or baudrate changed\n"
                              "• Test communication with new settings")
            
            # Reset connection tested flag
            if self.config_tool:
                self.config_tool.connection_tested = False
        else:
            self.log(f"❌ Configuration write failed: {message}", "error")
            messagebox.showerror("Write Failed", f"❌ Configuration write failed:\n\n{message}")
    
    # ==================== Logging ====================
    
    def log(self, message, tag=None):
        """Add message to activity log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, formatted_message, tag)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log(self):
        """Clear activity log"""
        self.log_text.delete(1.0, tk.END)
        self.log("🗑️ Activity log cleared", "info")
    
    def save_log(self):
        """Save activity log to file"""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfile=f"modbus_config_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )
            
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write("MODBUS RTU CONFIGURATION TOOL - ACTIVITY LOG\n")
                    f.write("=" * 60 + "\n")
                    f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("=" * 60 + "\n\n")
                    f.write(self.log_text.get(1.0, tk.END))
                
                messagebox.showinfo("Success", f"Log saved to:\n{filename}")
                self.log(f"💾 Log saved to {os.path.basename(filename)}", "success")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save log:\n{str(e)}")
            self.log(f"❌ Failed to save log: {e}", "error")
    
    # ==================== Window Management ====================
    
    def on_closing(self):
        """Handle window close event"""
        if self.is_connected:
            if messagebox.askokcancel("Quit", 
                                    "Connection is active. Disconnect and exit?"):
                self.disconnect()
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    """Main entry point"""
    try:
        root = tk.Tk()
        app = ModbusConfigGUI(root)
        
        # Center window on screen
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        # Add welcome message
        app.log("🚀 Modbus RTU Configuration Tool started", "success")
        app.log("📋 Configure connection settings and connect to device", "info")
        app.log("💡 Tip: Use 'Test Communication' before writing configuration", "info")
        
        root.mainloop()
        
    except Exception as e:
        print(f"Application error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()