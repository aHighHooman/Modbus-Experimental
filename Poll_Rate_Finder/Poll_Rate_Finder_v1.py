"""
RS485 Polling Rate Monitor - Refactored
=======================================

A comprehensive tool for monitoring and analyzing RS485 communication patterns,
including polling rates, wave detection, and protocol analysis.
This refactored version uses the ModbusMaster base class for improved maintainability.

Features:
- Passive RS485 monitoring
- Real-time polling rate analysis
- Wave/burst pattern detection
- Protocol detection (Modbus RTU/ASCII)
- Modern GUI with live statistics
- Cross-platform compatibility

Author: Your Team
Version: 2.0 (Refactored)
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import queue
from collections import defaultdict, deque
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any
import struct

# Import base Modbus class
try:
    from Templates.Modbus_Master_Base_v1 import ModbusMaster, SerialConfig, ModbusException
except ImportError:
    import tkinter.messagebox as mb
    mb.showerror("Missing Dependency", "modbus_master.py is required.\nPlease ensure it's in the same directory.")
    import sys
    sys.exit(1)


class RS485PassiveMonitor(ModbusMaster):
    """
    Specialized passive RS485 monitor for analyzing communication patterns.
    Extends ModbusMaster with passive listening capabilities.
    """
    
    def __init__(self, slave_ids: Optional[List[int]] = None, wave_gap_threshold: float = 0.5):
        super().__init__()
        
        # Configuration
        self.slave_ids = set(slave_ids) if slave_ids else set()
        self.wave_gap_threshold = wave_gap_threshold
        self.monitoring = False
        self.data_queue = queue.Queue()
        
        # Statistics tracking
        self.slave_stats = defaultdict(lambda: {
            'count': 0,
            'last_seen': None,
            'intervals': deque(maxlen=100),
            'avg_rate': 0,
            'min_interval': float('inf'),
            'max_interval': 0
        })
        
        # Wave/Burst detection
        self.wave_stats = defaultdict(lambda: {
            'waves': deque(maxlen=100),
            'current_wave_start': None,
            'current_wave_msgs': 0,
            'last_msg_time': None,
            'wave_intervals': deque(maxlen=100),
            'avg_wave_interval': 0,
            'avg_msgs_per_wave': 0,
            'min_wave_interval': float('inf'),
            'max_wave_interval': 0
        })
        
        # Protocol detection
        self.detected_protocols = {
            'modbus_rtu': False,
            'modbus_ascii': False,
            'custom': False
        }
    
    def setup_passive_mode(self, port: str, baudrate: int, timeout: float = 1.0):
        """
        Configure for passive RS485 monitoring
        
        Args:
            port: Serial port name
            baudrate: Communication baud rate
            timeout: Read timeout
        """
        config = SerialConfig(
            port=port,
            baudrate=baudrate,
            timeout=timeout,
            rtscts=False,  # No hardware flow control
            dsrdtr=False   # No DSR/DTR
        )
        
        if self.connect(config):
            # Configure for receive-only mode (high impedance)
            if self.serial_connection:
                self.serial_connection.rts = False
                self.serial_connection.dtr = False
            return True
        return False
    
    def auto_detect_baudrate(self, port: str) -> Optional[int]:
        """
        Auto-detect the communication baud rate by analyzing traffic patterns
        
        Args:
            port: Serial port to analyze
            
        Returns:
            Detected baud rate or None
        """
        print("🔍 Auto-detecting baud rate...")
        
        for baud in self.COMMON_BAUDRATES:
            print(f"Testing {baud:,} baud...", end=' ')
            
            if self.setup_passive_mode(port, baud, timeout=0.5):
                # Listen for valid frames
                valid_frames = 0
                start_time = time.time()
                
                while time.time() - start_time < 3:  # Test for 3 seconds
                    frames = self._capture_frames(max_frames=10, timeout=1.0)
                    
                    for frame in frames:
                        if self._analyze_frame_validity(frame):
                            valid_frames += 1
                
                self.disconnect()
                
                if valid_frames >= 2:  # Need at least 2 valid frames
                    print(f"✅ Success! ({valid_frames} valid frames)")
                    return baud
                else:
                    print(f"❌ No valid data")
            else:
                print("❌ Connection failed")
        
        print("❌ Could not auto-detect baud rate")
        return None
    
    def _capture_frames(self, max_frames: int = 100, timeout: float = 5.0) -> List[bytes]:
        """
        Capture raw frames from RS485 bus
        
        Args:
            max_frames: Maximum number of frames to capture
            timeout: Total capture timeout
            
        Returns:
            List of captured frame bytes
        """
        frames = []
        buffer = bytearray()
        
        # Calculate inter-frame timeout based on baud rate
        if self.config:
            char_time = 11.0 / self.config.baudrate
            frame_timeout = max(0.0015, char_time * 3.5)
        else:
            frame_timeout = 0.005
        
        start_time = time.time()
        last_byte_time = start_time
        
        while (time.time() - start_time < timeout and 
               len(frames) < max_frames and 
               self.serial_connection and 
               self.serial_connection.is_open):
            
            try:
                if self.serial_connection.in_waiting > 0:
                    chunk = self.serial_connection.read(self.serial_connection.in_waiting)
                    
                    if chunk:
                        # Check for frame gap
                        if buffer and (time.time() - last_byte_time > frame_timeout):
                            # Previous frame complete
                            if len(buffer) >= 4:  # Minimum frame size
                                frames.append(bytes(buffer))
                            buffer.clear()
                        
                        buffer.extend(chunk)
                        last_byte_time = time.time()
                
                elif buffer and (time.time() - last_byte_time > frame_timeout):
                    # Frame complete due to timeout
                    if len(buffer) >= 4:
                        frames.append(bytes(buffer))
                    buffer.clear()
                
                time.sleep(0.001)  # 1ms polling interval
                
            except Exception as e:
                print(f"Capture error: {e}")
                break
        
        # Add final buffer if it contains data
        if buffer and len(buffer) >= 4:
            frames.append(bytes(buffer))
        
        return frames
    
    def _analyze_frame_validity(self, frame: bytes) -> bool:
        """
        Analyze if a frame appears to be valid Modbus or similar protocol
        
        Args:
            frame: Raw frame bytes
            
        Returns:
            True if frame appears valid
        """
        if len(frame) < 4:
            return False
        
        # Check for Modbus RTU pattern
        slave_id = frame[0]
        function_code = frame[1]
        
        # Valid Modbus slave ID range
        if not (1 <= slave_id <= 247):
            return False
        
        # Common Modbus function codes
        valid_functions = [0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x0F, 0x10]
        exception_functions = [f | 0x80 for f in valid_functions]
        
        if function_code in valid_functions or function_code in exception_functions:
            # Verify CRC if possible
            if len(frame) >= 5:
                payload = frame[:-2]
                frame_crc = struct.unpack('<H', frame[-2:])[0]
                calculated_crc = self.calculate_crc(payload)
                
                if frame_crc == calculated_crc:
                    self.detected_protocols['modbus_rtu'] = True
                    return True
        
        # Check for Modbus ASCII
        if frame[0] == ord(':'):
            self.detected_protocols['modbus_ascii'] = True
            return True
        
        # Generic protocol detection
        if any(1 <= b <= 247 for b in frame):
            self.detected_protocols['custom'] = True
            return True
        
        return False
    
    def start_monitoring(self) -> bool:
        """
        Start passive monitoring of RS485 traffic
        
        Returns:
            True if monitoring started successfully
        """
        if not self.is_connected:
            return False
        
        self.monitoring = True
        monitor_thread = threading.Thread(target=self._monitor_thread, daemon=True)
        monitor_thread.start()
        return True
    
    def stop_monitoring(self):
        """Stop monitoring"""
        self.monitoring = False
    
    def _monitor_thread(self):
        """Main monitoring thread"""
        print("🎧 Starting passive RS485 monitoring...")
        
        while self.monitoring and self.is_connected:
            try:
                # Capture frames in batches
                frames = self._capture_frames(max_frames=50, timeout=1.0)
                
                for frame in frames:
                    if not self.monitoring:
                        break
                    self._process_frame(frame)
                
                time.sleep(0.1)  # Brief pause between batches
                
            except Exception as e:
                print(f"Monitor thread error: {e}")
                break
        
        print("🔇 Monitoring stopped")
    
    def _process_frame(self, frame: bytes):
        """
        Process a captured frame and update statistics
        
        Args:
            frame: Raw frame bytes
        """
        if not self._analyze_frame_validity(frame):
            return
        
        timestamp = datetime.now()
        slave_id = frame[0]
        
        # Auto-discover slave IDs
        if not self.slave_ids or slave_id in self.slave_ids:
            if slave_id not in self.slave_ids:
                self.slave_ids.add(slave_id)
                print(f"📡 Discovered new slave ID: {slave_id}")
            
            # Update wave detection
            self._update_wave_stats(slave_id, timestamp)
            
            # Update polling statistics
            stats = self.slave_stats[slave_id]
            min_poll_interval = 0.010  # 10ms minimum between legitimate polls
            
            if stats['last_seen']:
                interval = (timestamp - stats['last_seen']).total_seconds()
                
                # Only count as new poll if enough time has passed
                if interval >= min_poll_interval:
                    stats['intervals'].append(interval)
                    stats['min_interval'] = min(stats['min_interval'], interval)
                    stats['max_interval'] = max(stats['max_interval'], interval)
                    
                    # Calculate average rate
                    if stats['intervals']:
                        stats['avg_rate'] = sum(stats['intervals']) / len(stats['intervals'])
                    
                    stats['count'] += 1
                    stats['last_seen'] = timestamp
                    
                    # Queue for GUI display
                    self.data_queue.put({
                        'timestamp': timestamp,
                        'slave_id': slave_id,
                        'frame': frame.hex(),
                        'length': len(frame),
                        'interval': interval
                    })
            else:
                # First time seeing this slave
                stats['count'] += 1
                stats['last_seen'] = timestamp
                
                self.data_queue.put({
                    'timestamp': timestamp,
                    'slave_id': slave_id,
                    'frame': frame.hex(),
                    'length': len(frame),
                    'interval': 0
                })
    
    def _update_wave_stats(self, slave_id: int, timestamp: datetime):
        """
        Update wave/burst detection statistics
        
        Args:
            slave_id: Slave device ID
            timestamp: Frame timestamp
        """
        wave_stats = self.wave_stats[slave_id]
        
        if wave_stats['last_msg_time']:
            gap = (timestamp - wave_stats['last_msg_time']).total_seconds()
            
            # Check if this starts a new wave
            if gap > self.wave_gap_threshold:
                # Close previous wave
                if wave_stats['current_wave_start']:
                    wave_duration = (wave_stats['last_msg_time'] - wave_stats['current_wave_start']).total_seconds()
                    
                    wave_info = {
                        'start': wave_stats['current_wave_start'],
                        'end': wave_stats['last_msg_time'],
                        'duration': wave_duration,
                        'msg_count': wave_stats['current_wave_msgs']
                    }
                    wave_stats['waves'].append(wave_info)
                    
                    # Calculate wave intervals
                    if len(wave_stats['waves']) > 1:
                        prev_wave = wave_stats['waves'][-2]
                        wave_interval = (wave_stats['current_wave_start'] - prev_wave['start']).total_seconds()
                        wave_stats['wave_intervals'].append(wave_interval)
                        wave_stats['min_wave_interval'] = min(wave_stats['min_wave_interval'], wave_interval)
                        wave_stats['max_wave_interval'] = max(wave_stats['max_wave_interval'], wave_interval)
                    
                    # Update averages
                    if wave_stats['wave_intervals']:
                        wave_stats['avg_wave_interval'] = sum(wave_stats['wave_intervals']) / len(wave_stats['wave_intervals'])
                    
                    if wave_stats['waves']:
                        total_msgs = sum(w['msg_count'] for w in wave_stats['waves'])
                        wave_stats['avg_msgs_per_wave'] = total_msgs / len(wave_stats['waves'])
                
                # Start new wave
                wave_stats['current_wave_start'] = timestamp
                wave_stats['current_wave_msgs'] = 1
            else:
                # Continue current wave
                wave_stats['current_wave_msgs'] += 1
        else:
            # First message for this slave
            wave_stats['current_wave_start'] = timestamp
            wave_stats['current_wave_msgs'] = 1
        
        wave_stats['last_msg_time'] = timestamp
    
    def get_protocol_name(self) -> str:
        """Get detected protocol name"""
        if self.detected_protocols['modbus_rtu']:
            return "Modbus RTU"
        elif self.detected_protocols['modbus_ascii']:
            return "Modbus ASCII"
        elif self.detected_protocols['custom']:
            return "Custom/Unknown"
        return "Detecting..."
    
    def get_current_rates(self, window_seconds: float = 5.0) -> Dict[int, Dict[str, Any]]:
        """
        Get current polling rates for all slaves within a time window
        
        Args:
            window_seconds: Time window in seconds
            
        Returns:
            Dictionary of slave rates and statistics
        """
        current_time = datetime.now()
        rates = {}
        
        for slave_id in self.slave_ids:
            stats = self.slave_stats[slave_id]
            wave_stats = self.wave_stats[slave_id]
            
            # Calculate time since last message
            time_since_last = float('inf')
            if stats['last_seen']:
                time_since_last = (current_time - stats['last_seen']).total_seconds()
            
            # Determine status
            if time_since_last < 1.0:
                status = "🟢 Active"
            elif time_since_last < 3.0:
                status = "🟡 Slowing"
            else:
                status = "🔴 Idle"
            
            # Calculate current rate from recent intervals
            recent_intervals = [i for i in stats['intervals'] if i > 0]
            avg_rate = 1.0 / (sum(recent_intervals) / len(recent_intervals)) if recent_intervals else 0
            
            rates[slave_id] = {
                'rate_hz': avg_rate,
                'period_ms': (sum(recent_intervals) / len(recent_intervals) * 1000) if recent_intervals else 0,
                'message_count': stats['count'],
                'status': status,
                'time_since_last': time_since_last,
                'current_wave_msgs': wave_stats['current_wave_msgs'],
                'total_waves': len(wave_stats['waves']),
                'avg_msgs_per_wave': wave_stats['avg_msgs_per_wave']
            }
        
        return rates


class WaveAnalysisGUI:
    """
    Modern GUI for RS485 wave pattern analysis with real-time monitoring
    """
    
    def __init__(self, root):
        self.root = root
        self.root.title("RS485 Polling Rate & Wave Analysis Monitor")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)
        
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
            'light': '#f8f9fa'
        }
        
        self.root.configure(bg=self.colors['bg'])
        
        # Initialize monitor and variables
        self.monitor: Optional[RS485PassiveMonitor] = None
        self.monitor_thread: Optional[threading.Thread] = None
        self.update_queue = queue.Queue()
        self.monitoring = False
        
        # GUI variables
        self.port_var = tk.StringVar()
        self.baud_var = tk.StringVar(value="19200")
        self.wave_gap_var = tk.StringVar(value="0.5")
        self.slaves_var = tk.StringVar(value="")
        self.update_rate_var = tk.StringVar(value="1.0")
        
        # Create UI
        self.setup_styles()
        self.create_widgets()
        self.refresh_ports()
        self.start_update_loop()
    
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
        style.configure("Title.TLabel", font=('Segoe UI', 16, 'bold'), foreground=self.colors['dark'])
        style.configure("Header.TLabel", font=('Segoe UI', 11, 'bold'), foreground=self.colors['dark'])
        style.configure("Info.TLabel", font=('Segoe UI', 9), foreground=self.colors['info'])
        style.configure("Success.TButton", foreground='white')
        style.configure("Danger.TButton", foreground='white')
    
    def create_widgets(self):
        """Create all GUI widgets"""
        # Main container with proper grid configuration
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Create sections
        self.create_title_section(main_frame)
        self.create_config_panel(main_frame)
        self.create_display_area(main_frame)
        self.create_status_bar()
    
    def create_title_section(self, parent):
        """Create title section"""
        title_frame = tk.Frame(parent, bg=self.colors['bg'])
        title_frame.grid(row=0, column=0, sticky='ew', pady=(0, 20))
        
        ttk.Label(title_frame, text="🌊 RS485 Polling Rate & Wave Analysis Monitor", 
                 style="Title.TLabel").pack()
        ttk.Label(title_frame, text="Real-time monitoring and analysis of RS485 communication patterns", 
                 style="Info.TLabel").pack(pady=(5, 0))
    
    def create_config_panel(self, parent):
        """Create configuration panel"""
        config_frame = ttk.LabelFrame(parent, text="Configuration", padding="15")
        config_frame.grid(row=1, column=0, sticky='ew', pady=(0, 15))
        
        # Connection settings row
        conn_row = ttk.Frame(config_frame)
        conn_row.pack(fill=tk.X, pady=(0, 10))
        
        # Port selection
        ttk.Label(conn_row, text="Port:", font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(0, 5))
        self.port_combo = ttk.Combobox(conn_row, textvariable=self.port_var, 
                                      width=15, state='readonly')
        self.port_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(conn_row, text="🔄", width=3, command=self.refresh_ports).pack(side=tk.LEFT, padx=(0, 20))
        
        # Baudrate
        ttk.Label(conn_row, text="Baudrate:", font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(0, 5))
        baud_combo = ttk.Combobox(conn_row, textvariable=self.baud_var,
                                 values=["1200", "2400", "4800", "9600", "19200", "38400", "57600", "115200"],
                                 width=10, state='readonly')
        baud_combo.pack(side=tk.LEFT, padx=(0, 20))
        
        # Auto-detect button
        ttk.Button(conn_row, text="🔍 Auto-detect", command=self.auto_detect_baudrate).pack(side=tk.LEFT)
        
        # Analysis settings row
        analysis_row = ttk.Frame(config_frame)
        analysis_row.pack(fill=tk.X, pady=(0, 10))
        
        # Wave gap threshold
        ttk.Label(analysis_row, text="Wave Gap (s):", font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(0, 5))
        wave_spin = tk.Spinbox(analysis_row, from_=0.1, to=5.0, increment=0.1, 
                              width=8, textvariable=self.wave_gap_var)
        wave_spin.pack(side=tk.LEFT, padx=(0, 20))
        
        # Slave IDs (optional)
        ttk.Label(analysis_row, text="Monitor Slaves (optional):", font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(analysis_row, textvariable=self.slaves_var, width=15).pack(side=tk.LEFT, padx=(0, 20))
        
        # Update rate
        ttk.Label(analysis_row, text="Update Rate (s):", font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(0, 5))
        update_spin = tk.Spinbox(analysis_row, from_=0.5, to=5.0, increment=0.5, 
                                width=8, textvariable=self.update_rate_var)
        update_spin.pack(side=tk.LEFT)
        
        # Control buttons row
        control_row = ttk.Frame(config_frame)
        control_row.pack(fill=tk.X)
        
        self.start_btn = ttk.Button(control_row, text="▶ Start Monitoring", 
                                   command=self.start_monitoring, style="Header.TLabel")
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_btn = ttk.Button(control_row, text="⏹ Stop", 
                                  command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(control_row, text="🗑 Clear Display", 
                  command=self.clear_display).pack(side=tk.LEFT)
    
    def create_display_area(self, parent):
        """Create main display area with notebook"""
        notebook = ttk.Notebook(parent)
        notebook.grid(row=2, column=0, sticky='nsew')
        
        # Create tabs
        self.create_rates_tab(notebook)
        self.create_waves_tab(notebook)
        self.create_stats_tab(notebook)
    
    def create_rates_tab(self, notebook):
        """Create real-time rates monitoring tab"""
        rates_frame = ttk.Frame(notebook)
        notebook.add(rates_frame, text="📊 Real-time Rates")
        rates_frame.columnconfigure(0, weight=1)
        rates_frame.rowconfigure(1, weight=1)
        
        # Summary cards
        cards_frame = tk.Frame(rates_frame, bg=self.colors['bg'])
        cards_frame.grid(row=0, column=0, sticky='ew', padx=10, pady=10)
        
        self.cards = {}
        card_configs = [
            ('Active Slaves', '0', self.colors['success']),
            ('Total Messages', '0', self.colors['info']),
            ('Avg Rate', '0.00 Hz', self.colors['warning']),
            ('Protocol', 'Detecting...', self.colors['accent'])
        ]
        
        for i, (label, value, color) in enumerate(card_configs):
            card = self.create_info_card(cards_frame, label, value, color)
            card['frame'].pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
            self.cards[label] = card['value_label']
        
        # Rates display
        rates_display = tk.Frame(rates_frame, bg='white', relief=tk.SUNKEN, bd=1)
        rates_display.grid(row=1, column=0, sticky='nsew', padx=10, pady=(0, 10))
        rates_display.columnconfigure(0, weight=1)
        rates_display.rowconfigure(0, weight=1)
        
        self.rates_text = tk.Text(rates_display, wrap=tk.NONE,
                                 bg='#1a1a1a', fg='#00ff00',
                                 font=('Consolas', 10),
                                 relief=tk.FLAT, bd=0)
        
        rates_scroll_y = tk.Scrollbar(rates_display, orient=tk.VERTICAL, command=self.rates_text.yview)
        rates_scroll_x = tk.Scrollbar(rates_display, orient=tk.HORIZONTAL, command=self.rates_text.xview)
        
        self.rates_text.configure(yscrollcommand=rates_scroll_y.set, xscrollcommand=rates_scroll_x.set)
        
        self.rates_text.grid(row=0, column=0, sticky='nsew')
        rates_scroll_y.grid(row=0, column=1, sticky='ns')
        rates_scroll_x.grid(row=1, column=0, sticky='ew')
    
    def create_waves_tab(self, notebook):
        """Create wave analysis tab"""
        waves_frame = ttk.Frame(notebook)
        notebook.add(waves_frame, text="🌊 Wave Analysis")
        waves_frame.columnconfigure(0, weight=1)
        waves_frame.rowconfigure(0, weight=1)
        
        # Wave display
        waves_display = tk.Frame(waves_frame, bg='white', relief=tk.SUNKEN, bd=1)
        waves_display.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
        waves_display.columnconfigure(0, weight=1)
        waves_display.rowconfigure(0, weight=1)
        
        self.waves_text = tk.Text(waves_display, wrap=tk.NONE,
                                 bg='#1a1a2e', fg='#00ffff',
                                 font=('Consolas', 10),
                                 relief=tk.FLAT, bd=0)
        
        waves_scroll_y = tk.Scrollbar(waves_display, orient=tk.VERTICAL, command=self.waves_text.yview)
        waves_scroll_x = tk.Scrollbar(waves_display, orient=tk.HORIZONTAL, command=self.waves_text.xview)
        
        self.waves_text.configure(yscrollcommand=waves_scroll_y.set, xscrollcommand=waves_scroll_x.set)
        
        self.waves_text.grid(row=0, column=0, sticky='nsew')
        waves_scroll_y.grid(row=0, column=1, sticky='ns')
        waves_scroll_x.grid(row=1, column=0, sticky='ew')
    
    def create_stats_tab(self, notebook):
        """Create detailed statistics tab"""
        stats_frame = ttk.Frame(notebook)
        notebook.add(stats_frame, text="📈 Statistics")
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)
        
        # Statistics display
        stats_display = tk.Frame(stats_frame, bg='white', relief=tk.SUNKEN, bd=1)
        stats_display.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
        stats_display.columnconfigure(0, weight=1)
        stats_display.rowconfigure(0, weight=1)
        
        self.stats_text = tk.Text(stats_display, wrap=tk.NONE,
                                 bg='#2d1b69', fg='#ffffff',
                                 font=('Consolas', 10),
                                 relief=tk.FLAT, bd=0)
        
        stats_scroll_y = tk.Scrollbar(stats_display, orient=tk.VERTICAL, command=self.stats_text.yview)
        stats_scroll_x = tk.Scrollbar(stats_display, orient=tk.HORIZONTAL, command=self.stats_text.xview)
        
        self.stats_text.configure(yscrollcommand=stats_scroll_y.set, xscrollcommand=stats_scroll_x.set)
        
        self.stats_text.grid(row=0, column=0, sticky='nsew')
        stats_scroll_y.grid(row=0, column=1, sticky='ns')
        stats_scroll_x.grid(row=1, column=0, sticky='ew')
    
    def create_info_card(self, parent, label, value, color):
        """Create an information display card"""
        card_frame = tk.Frame(parent, bg='white', relief=tk.RAISED, bd=1)
        
        label_widget = tk.Label(card_frame, text=label, bg='white', fg='#666666', 
                               font=('Segoe UI', 9))
        label_widget.pack(pady=(10, 5))
        
        value_label = tk.Label(card_frame, text=value, bg='white', fg=color, 
                             font=('Segoe UI', 14, 'bold'))
        value_label.pack(pady=(0, 10))
        
        return {'frame': card_frame, 'value_label': value_label}
    
    def create_status_bar(self):
        """Create status bar"""
        status_frame = tk.Frame(self.root, bg='#e9ecef', relief=tk.SUNKEN, bd=1)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_label = tk.Label(status_frame, text="Ready - Configure settings and start monitoring", 
                                   bg='#e9ecef', fg='#495057', anchor='w')
        self.status_label.pack(side=tk.LEFT, padx=10, pady=3)
        
        self.connection_label = tk.Label(status_frame, text="⚫ Disconnected", 
                                       bg='#e9ecef', fg=self.colors['danger'], anchor='e')
        self.connection_label.pack(side=tk.RIGHT, padx=10, pady=3)
    
    # ==================== Port Management ====================
    
    def refresh_ports(self):
        """Refresh available COM ports"""
        try:
            ports = ModbusMaster.list_available_ports()
            port_list = [p['port'] for p in ports]
            
            self.port_combo['values'] = port_list
            if port_list:
                self.port_combo.current(0)
            
            self.status_label.config(text=f"Found {len(ports)} serial ports")
            
        except Exception as e:
            self.status_label.config(text=f"Error refreshing ports: {e}")
    
    def auto_detect_baudrate(self):
        """Auto-detect baud rate"""
        port = self.port_var.get()
        if not port:
            messagebox.showerror("Error", "Please select a port first")
            return
        
        self.status_label.config(text="Auto-detecting baud rate...")
        self.start_btn.config(state=tk.DISABLED)
        
        def detect_worker():
            try:
                monitor = RS485PassiveMonitor()
                detected_baud = monitor.auto_detect_baudrate(port)
                self.root.after(0, lambda: self._auto_detect_complete(detected_baud))
            except Exception as e:
                self.root.after(0, lambda: self._auto_detect_complete(None, str(e)))
        
        threading.Thread(target=detect_worker, daemon=True).start()
    
    def _auto_detect_complete(self, detected_baud, error=None):
        """Handle auto-detection completion"""
        self.start_btn.config(state=tk.NORMAL)
        
        if detected_baud and not error:
            self.baud_var.set(str(detected_baud))
            messagebox.showinfo("Success", f"Detected baud rate: {detected_baud:,}")
            self.status_label.config(text=f"Auto-detected baud rate: {detected_baud:,}")
        else:
            error_msg = error if error else "Could not auto-detect baud rate"
            messagebox.showerror("Failed", error_msg)
            self.status_label.config(text="Auto-detection failed")
    
    # ==================== Monitoring Control ====================
    
    def start_monitoring(self):
        """Start RS485 monitoring"""
        if self.monitoring:
            return
        
        port = self.port_var.get()
        if not port:
            messagebox.showerror("Error", "Please select a port")
            return
        
        try:
            baudrate = int(self.baud_var.get())
            wave_gap = float(self.wave_gap_var.get())
            
            # Parse slave IDs if provided
            slaves = None
            if self.slaves_var.get().strip():
                try:
                    slaves = [int(x.strip()) for x in self.slaves_var.get().split(',') if x.strip()]
                except ValueError:
                    messagebox.showerror("Error", "Invalid slave ID format. Use comma-separated numbers.")
                    return
            
            # Create monitor
            self.monitor = RS485PassiveMonitor(slave_ids=slaves, wave_gap_threshold=wave_gap)
            
            # Setup and connect
            if not self.monitor.setup_passive_mode(port, baudrate, timeout=1.0):
                messagebox.showerror("Error", f"Failed to connect to {port}")
                return
            
            # Start monitoring
            if not self.monitor.start_monitoring():
                messagebox.showerror("Error", "Failed to start monitoring")
                return
            
            # Update UI
            self.monitoring = True
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.connection_label.config(text=f"🟢 Connected to {port}", fg=self.colors['success'])
            self.status_label.config(text=f"Monitoring at {baudrate:,} baud")
            
        except ValueError:
            messagebox.showerror("Error", "Invalid baud rate or wave gap value")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start monitoring: {str(e)}")
    
    def stop_monitoring(self):
        """Stop monitoring"""
        if not self.monitoring:
            return
        
        self.monitoring = False
        
        if self.monitor:
            self.monitor.stop_monitoring()
            self.monitor.disconnect()
            self.monitor = None
        
        # Update UI
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.connection_label.config(text="⚫ Disconnected", fg=self.colors['danger'])
        self.status_label.config(text="Monitoring stopped")
    
    def clear_display(self):
        """Clear all displays"""
        self.rates_text.delete(1.0, tk.END)
        self.waves_text.delete(1.0, tk.END)
        self.stats_text.delete(1.0, tk.END)
        
        # Reset cards
        for card in self.cards.values():
            card.config(text="0")
    
    # ==================== Display Updates ====================
    
    def start_update_loop(self):
        """Start the display update loop"""
        self.update_display()
    
    def update_display(self):
        """Update all displays with current data"""
        if self.monitoring and self.monitor:
            try:
                # Get current rates and statistics
                rates = self.monitor.get_current_rates()
                protocol = self.monitor.get_protocol_name()
                
                # Update summary cards
                active_slaves = sum(1 for r in rates.values() if "Active" in r['status'])
                total_messages = sum(r['message_count'] for r in rates.values())
                avg_rate = sum(r['rate_hz'] for r in rates.values()) / len(rates) if rates else 0
                
                self.cards['Active Slaves'].config(text=str(active_slaves))
                self.cards['Total Messages'].config(text=f"{total_messages:,}")
                self.cards['Avg Rate'].config(text=f"{avg_rate:.2f} Hz")
                self.cards['Protocol'].config(text=protocol)
                
                # Update displays
                self.update_rates_display(rates)
                self.update_waves_display()
                self.update_stats_display()
                
            except Exception as e:
                print(f"Display update error: {e}")
        
        # Schedule next update
        try:
            update_rate = float(self.update_rate_var.get()) * 1000  # Convert to milliseconds
            self.root.after(int(update_rate), self.update_display)
        except (ValueError, tk.TclError):
            # If there's an error, default to 1 second
            self.root.after(1000, self.update_display)
    
    def update_rates_display(self, rates):
        """Update real-time rates display"""
        self.rates_text.delete(1.0, tk.END)
        
        # Header
        self.rates_text.insert(tk.END, "=" * 80 + "\n")
        self.rates_text.insert(tk.END, f"RS485 REAL-TIME POLLING RATES - {datetime.now().strftime('%H:%M:%S')}\n")
        self.rates_text.insert(tk.END, "=" * 80 + "\n\n")
        
        if not rates:
            self.rates_text.insert(tk.END, "Waiting for data...\n")
            self.rates_text.insert(tk.END, "Check connections if no data appears after 30 seconds.\n")
            return
        
        # Table header
        self.rates_text.insert(tk.END, f"{'Slave ID':<12} {'Rate (Hz)':<12} {'Period (ms)':<12} "
                                      f"{'Messages':<12} {'Status':<15} {'Wave Msgs':<12}\n")
        self.rates_text.insert(tk.END, "-" * 80 + "\n")
        
        # Display rates for each slave
        for slave_id in sorted(rates.keys()):
            data = rates[slave_id]
            
            self.rates_text.insert(tk.END, 
                f"{slave_id:<12} {data['rate_hz']:<12.2f} {data['period_ms']:<12.1f} "
                f"{data['message_count']:<12} {data['status']:<15} {data['current_wave_msgs']:<12}\n"
            )
        
        # Summary
        total_rate = sum(r['rate_hz'] for r in rates.values())
        self.rates_text.insert(tk.END, "\n" + "-" * 80 + "\n")
        self.rates_text.insert(tk.END, f"Summary: {len(rates)} active slaves | Total rate: {total_rate:.1f} Hz\n")
        self.rates_text.insert(tk.END, "=" * 80 + "\n")
    
    def update_waves_display(self):
        """Update wave analysis display"""
        if not self.monitor:
            return
        
        self.waves_text.delete(1.0, tk.END)
        
        # Header
        self.waves_text.insert(tk.END, "=" * 90 + "\n")
        self.waves_text.insert(tk.END, f"WAVE/BURST PATTERN ANALYSIS - {datetime.now().strftime('%H:%M:%S')}\n")
        self.waves_text.insert(tk.END, f"Gap Threshold: {self.monitor.wave_gap_threshold}s | Protocol: {self.monitor.get_protocol_name()}\n")
        self.waves_text.insert(tk.END, "=" * 90 + "\n\n")
        
        if not self.monitor.wave_stats:
            self.waves_text.insert(tk.END, "Waiting for wave data...\n")
            return
        
        # Current wave activity
        self.waves_text.insert(tk.END, "CURRENT WAVE ACTIVITY:\n")
        self.waves_text.insert(tk.END, "-" * 90 + "\n")
        self.waves_text.insert(tk.END, f"{'Slave ID':<12} {'Status':<20} {'Current Msgs':<15} {'Time Since Last':<20}\n")
        self.waves_text.insert(tk.END, "-" * 90 + "\n")
        
        current_time = datetime.now()
        for slave_id in sorted(self.monitor.wave_stats.keys()):
            wave_stat = self.monitor.wave_stats[slave_id]
            
            if wave_stat['last_msg_time']:
                time_since = (current_time - wave_stat['last_msg_time']).total_seconds()
                
                if time_since < self.monitor.wave_gap_threshold:
                    status = "🟢 In Wave"
                elif time_since < self.monitor.wave_gap_threshold * 3:
                    status = "🟡 Wave Ending"
                else:
                    status = "🔴 Between Waves"
                
                self.waves_text.insert(tk.END, 
                    f"{slave_id:<12} {status:<20} {wave_stat['current_wave_msgs']:<15} {time_since:.3f}s\n"
                )
        
        # Wave pattern statistics
        self.waves_text.insert(tk.END, "\nWAVE PATTERN STATISTICS:\n")
        self.waves_text.insert(tk.END, "-" * 90 + "\n")
        self.waves_text.insert(tk.END, f"{'Slave ID':<12} {'Waves':<12} {'Avg Msgs':<12} {'Wave Period':<15} "
                                      f"{'Frequency':<12} {'Range':<20}\n")
        self.waves_text.insert(tk.END, "-" * 90 + "\n")
        
        for slave_id in sorted(self.monitor.wave_stats.keys()):
            wave_stat = self.monitor.wave_stats[slave_id]
            waves_count = len(wave_stat['waves'])
            
            if waves_count > 0:
                period = wave_stat['avg_wave_interval']
                freq = 1.0 / period if period > 0 else 0
                period_range = f"{wave_stat['min_wave_interval']:.1f}-{wave_stat['max_wave_interval']:.1f}s"
                
                self.waves_text.insert(tk.END,
                    f"{slave_id:<12} {waves_count:<12} {wave_stat['avg_msgs_per_wave']:<12.1f} "
                    f"{period:<15.3f} {freq:<12.3f} {period_range:<20}\n"
                )
        
        # Recent wave history
        self.waves_text.insert(tk.END, "\nRECENT WAVE HISTORY (Last 5 waves per slave):\n")
        self.waves_text.insert(tk.END, "-" * 90 + "\n")
        
        for slave_id in sorted(self.monitor.wave_stats.keys()):
            wave_stat = self.monitor.wave_stats[slave_id]
            if wave_stat['waves']:
                recent_waves = list(wave_stat['waves'])[-5:]
                wave_summary = []
                for wave in recent_waves:
                    wave_summary.append(f"{wave['msg_count']}msgs/{wave['duration']:.2f}s")
                
                self.waves_text.insert(tk.END, f"Slave {slave_id:3}: {' | '.join(wave_summary)}\n")
        
        self.waves_text.insert(tk.END, "=" * 90 + "\n")
    
    def update_stats_display(self):
        """Update detailed statistics display"""
        if not self.monitor:
            return
        
        self.stats_text.delete(1.0, tk.END)
        
        # Header
        self.stats_text.insert(tk.END, "=" * 100 + "\n")
        self.stats_text.insert(tk.END, f"DETAILED COMMUNICATION STATISTICS - {datetime.now().strftime('%H:%M:%S')}\n")
        self.stats_text.insert(tk.END, "=" * 100 + "\n\n")
        
        # Individual message statistics
        self.stats_text.insert(tk.END, "INDIVIDUAL MESSAGE STATISTICS:\n")
        self.stats_text.insert(tk.END, "-" * 100 + "\n")
        self.stats_text.insert(tk.END, f"{'Slave ID':<10} {'Count':<10} {'Avg Interval':<15} "
                                      f"{'Min Interval':<15} {'Max Interval':<15} {'Frequency':<15}\n")
        self.stats_text.insert(tk.END, "-" * 100 + "\n")
        
        for slave_id in sorted(self.monitor.slave_ids):
            stats = self.monitor.slave_stats[slave_id]
            if stats['count'] > 1:
                freq = 1.0 / stats['avg_rate'] if stats['avg_rate'] > 0 else 0
                
                self.stats_text.insert(tk.END,
                    f"{slave_id:<10} {stats['count']:<10} {stats['avg_rate']:<15.3f} "
                    f"{stats['min_interval']:<15.3f} {stats['max_interval']:<15.3f} {freq:<15.2f}\n"
                )
        
        # Protocol detection results
        self.stats_text.insert(tk.END, "\nPROTOCOL DETECTION:\n")
        self.stats_text.insert(tk.END, "-" * 100 + "\n")
        for protocol, detected in self.monitor.detected_protocols.items():
            status = "✅ Detected" if detected else "❌ Not detected"
            self.stats_text.insert(tk.END, f"{protocol.replace('_', ' ').title():<20}: {status}\n")
        
        # Communication quality metrics
        if hasattr(self.monitor, 'get_statistics'):
            comm_stats = self.monitor.get_statistics()
            self.stats_text.insert(tk.END, "\nCOMMUNICATION QUALITY:\n")
            self.stats_text.insert(tk.END, "-" * 100 + "\n")
            self.stats_text.insert(tk.END, f"Success Rate: {comm_stats.get('success_rate', 0):.1f}%\n")
            self.stats_text.insert(tk.END, f"Error Rate: {comm_stats.get('error_rate', 0):.1f}%\n")
            self.stats_text.insert(tk.END, f"Total Requests: {comm_stats.get('requests_sent', 0)}\n")
            self.stats_text.insert(tk.END, f"Responses Received: {comm_stats.get('responses_received', 0)}\n")
            self.stats_text.insert(tk.END, f"Timeouts: {comm_stats.get('timeouts', 0)}\n")
            self.stats_text.insert(tk.END, f"CRC Errors: {comm_stats.get('crc_errors', 0)}\n")
        
        self.stats_text.insert(tk.END, "=" * 100 + "\n")


def main():
    """Main entry point"""
    try:
        root = tk.Tk()
        app = WaveAnalysisGUI(root)
        
        # Handle window closing
        def on_closing():
            if app.monitoring:
                app.stop_monitoring()
                time.sleep(0.5)
            root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()
        
    except Exception as e:
        print(f"Application error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()