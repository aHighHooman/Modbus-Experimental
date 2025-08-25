"""
Modbus RTU Master Base Class
===========================
Features:
- Serial communication
- CRC16 validation
- Port detection
- Device scanning and discovery
- Error handling and timeout management
- Thread-safe operations

Author: Umair
Version: 1.0
"""

import sys
import struct
import time
import threading
from typing import List, Tuple, Optional, Union, Dict, Any
from datetime import datetime
from dataclasses import dataclass
from enum import Enum
import logging

# Platform-specific imports
try:
    import serial
    import serial.tools.list_ports
except ImportError as e:
    print(f"Error: pyserial is required but not installed.")
    print("Please install it with: pip install pyserial")
    sys.exit(1)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ModbusException(Exception):
    """Base exception for Modbus operations"""
    pass


class ModbusTimeoutException(ModbusException):
    """Raised when a Modbus operation times out"""
    pass


class ModbusCRCException(ModbusException):
    """Raised when CRC validation fails"""
    pass


class FunctionCode(Enum):
    """Modbus function codes"""
    READ_COILS = 0x01
    READ_DISCRETE_INPUTS = 0x02
    READ_HOLDING_REGISTERS = 0x03
    READ_INPUT_REGISTERS = 0x04
    WRITE_SINGLE_COIL = 0x05
    WRITE_SINGLE_REGISTER = 0x06
    WRITE_MULTIPLE_COILS = 0x0F
    WRITE_MULTIPLE_REGISTERS = 0x10


@dataclass
class ModbusResponse:
    """Represents a Modbus response with all relevant information"""
    slave_id: Optional[int] = None
    function_code: Optional[int] = None
    data: Optional[bytes] = None
    exception_code: Optional[int] = None
    timeout: bool = False
    crc_error: bool = False
    raw_frame: Optional[bytes] = None
    timestamp: Optional[datetime] = None
    
    def __post_init__(self):
        if self.timestamp is None:
            self.timestamp = datetime.now()
    
    @property
    def is_valid(self) -> bool:
        """Check if response is valid (no timeout or CRC error)"""
        return not self.timeout and not self.crc_error
    
    @property
    def is_exception(self) -> bool:
        """Check if response is an exception response"""
        return self.exception_code is not None
    
    @property
    def exception_name(self) -> str:
        """Get human-readable exception name"""
        if self.exception_code is None:
            return "Normal Response"
        
        exceptions = {
            1: "Illegal Function",
            2: "Illegal Data Address", 
            3: "Illegal Data Value",
            4: "Slave Device Failure",
            5: "Acknowledge",
            6: "Slave Device Busy",
            8: "Memory Parity Error",
            10: "Gateway Path Unavailable",
            11: "Gateway Target Failed"
        }
        return exceptions.get(self.exception_code, f"Unknown Exception ({self.exception_code})")


@dataclass
class SerialConfig:
    """Serial port configuration"""
    port: str
    baudrate: int = 9600
    bytesize: int = serial.EIGHTBITS
    parity: str = serial.PARITY_NONE
    stopbits: int = serial.STOPBITS_ONE
    timeout: float = 2.0
    write_timeout: float = 2.0
    xonxoff: bool = False
    rtscts: bool = False
    dsrdtr: bool = False


class ModbusMaster:
    """
    Base Modbus RTU Master class providing comprehensive functionality
    for device communication, scanning, and configuration.
    """
    
    # Common baud rates for auto-detection
    COMMON_BAUDRATES = [9600, 19200, 38400, 57600, 115200, 4800, 2400, 1200]
    
    # Data format options for auto-detection
    DATA_FORMATS = [
        (8, 'N', 1),  # 8N1 (most common)
        (8, 'E', 1),  # 8E1
        (8, 'O', 1),  # 8O1
        (7, 'E', 1),  # 7E1
        (7, 'O', 1),  # 7O1
    ]
    
    def __init__(self, config: Optional[SerialConfig] = None):
        """
        Initialize Modbus Master
        
        Args:
            config: Serial configuration. If None, will need to be set later.
        """
        self.config = config
        self.serial_connection: Optional[serial.Serial] = None
        self.is_connected = False
        self._lock = threading.Lock()
        
        # Statistics and monitoring
        self.stats = {
            'requests_sent': 0,
            'responses_received': 0,
            'timeouts': 0,
            'crc_errors': 0,
            'exceptions': 0
        }
    
    # ==================== Connection Management ====================
    
    def connect(self, config: Optional[SerialConfig] = None) -> bool:
        """
        Connect to the Modbus device
        
        Args:
            config: Optional serial configuration to use
            
        Returns:
            True if connection successful, False otherwise
        """
        if config:
            self.config = config
            
        if not self.config:
            raise ModbusException("No serial configuration provided")
        
        try:
            with self._lock:
                # Close existing connection if open
                if self.serial_connection and self.serial_connection.is_open:
                    self.serial_connection.close()
                    time.sleep(0.1)
                
                # Create new connection
                self.serial_connection = serial.Serial(
                    port=self.config.port,
                    baudrate=self.config.baudrate,
                    bytesize=self.config.bytesize,
                    parity=self.config.parity,
                    stopbits=self.config.stopbits,
                    timeout=self.config.timeout,
                    write_timeout=self.config.write_timeout,
                    xonxoff=self.config.xonxoff,
                    rtscts=self.config.rtscts,
                    dsrdtr=self.config.dsrdtr
                )
                
                # Allow time for port to stabilize
                time.sleep(0.1)
                
                # Clear buffers
                self.serial_connection.reset_input_buffer()
                self.serial_connection.reset_output_buffer()
                
                self.is_connected = True
                logger.info(f"Connected to {self.config.port} at {self.config.baudrate} baud")
                return True
                
        except (serial.SerialException, OSError) as e:
            logger.error(f"Failed to connect to {self.config.port}: {e}")
            self.is_connected = False
            return False
    
    def disconnect(self) -> None:
        """Disconnect from the Modbus device"""
        try:
            with self._lock:
                if self.serial_connection and self.serial_connection.is_open:
                    self.serial_connection.close()
                    logger.info(f"Disconnected from {self.config.port}")
                self.is_connected = False
        except Exception as e:
            logger.error(f"Error during disconnect: {e}")
    
    def __enter__(self):
        """Context manager entry"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.disconnect()
    
    # ==================== Port Discovery ====================
    
    @staticmethod
    def list_available_ports() -> List[Dict[str, str]]:
        """
        List all available serial ports
        
        Returns:
            List of dictionaries with port information
        """
        ports = []
        for port_info in serial.tools.list_ports.comports():
            ports.append({
                'port': port_info.device,
                'description': port_info.description,
                'hwid': port_info.hwid,
                'manufacturer': getattr(port_info, 'manufacturer', 'Unknown')
            })
        return ports
    
    @staticmethod
    def find_rs485_ports() -> List[Dict[str, str]]:
        """
        Find ports that might be RS485 adapters
        
        Returns:
            List of potential RS485 adapter ports
        """
        rs485_keywords = ['RS485', 'RS-485', 'USB-485', 'Serial', 'FTDI', 'CH340', 'CP210']
        potential_ports = []
        
        for port_info in ModbusMaster.list_available_ports():
            description = port_info['description'].upper()
            if any(keyword.upper() in description for keyword in rs485_keywords):
                potential_ports.append(port_info)
        
        return potential_ports
    
    def auto_detect_port(self) -> Optional[str]:
        """
        Auto-detect RS485 adapter port
        
        Returns:
            Port name if found, None otherwise
        """
        rs485_ports = self.find_rs485_ports()
        if rs485_ports:
            logger.info(f"Found potential RS485 adapter: {rs485_ports[0]['port']}")
            return rs485_ports[0]['port']
        
        # Fallback to first available port
        all_ports = self.list_available_ports()
        if all_ports:
            logger.warning(f"No RS485 adapter found, using first available: {all_ports[0]['port']}")
            return all_ports[0]['port']
        
        logger.error("No serial ports available")
        return None
    
    # ==================== CRC Calculation ====================
    
    @staticmethod
    def calculate_crc(data: bytes) -> int:
        """
        Calculate Modbus CRC16
        
        Args:
            data: Data bytes to calculate CRC for
            
        Returns:
            CRC16 value
        """
        crc = 0xFFFF
        for byte in data:
            crc ^= byte
            for _ in range(8):
                if crc & 0x0001:
                    crc = (crc >> 1) ^ 0xA001
                else:
                    crc >>= 1
        return crc & 0xFFFF
    
    # ==================== Low-level Communication ====================
    
    def _send_frame(self, frame: bytes) -> ModbusResponse:
        """
        Send a Modbus frame and receive response
        
        Args:
            frame: Complete Modbus frame with CRC
            
        Returns:
            ModbusResponse object
        """
        if not self.is_connected or not self.serial_connection:
            return ModbusResponse(timeout=True)
        
        try:
            with self._lock:
                # Clear input buffer and send frame
                self.serial_connection.reset_input_buffer()
                time.sleep(0.005)  # Small delay for buffer clearing
                
                self.serial_connection.write(frame)
                self.serial_connection.flush()
                self.stats['requests_sent'] += 1
                
                # Read response with timeout handling
                response_data = self._read_response()
                
                if not response_data:
                    self.stats['timeouts'] += 1
                    return ModbusResponse(timeout=True)
                
                # Parse the response
                return self._parse_response(response_data)
                
        except Exception as e:
            logger.error(f"Communication error: {e}")
            self.stats['timeouts'] += 1
            return ModbusResponse(timeout=True)
    
    def _read_response(self) -> Optional[bytes]:
        """
        Read response from serial port with proper frame detection
        
        Returns:
            Complete response frame or None if timeout
        """
        buffer = bytearray()
        start_time = time.time()
        last_byte_time = start_time
        
        # Calculate inter-frame timeout based on baudrate
        char_time = 11.0 / self.config.baudrate  # 11 bits per character
        frame_timeout = max(0.0015, char_time * 3.5)  # 3.5 character times or 1.5ms min
        
        while time.time() - start_time < self.config.timeout:
            if self.serial_connection.in_waiting > 0:
                chunk = self.serial_connection.read(self.serial_connection.in_waiting)
                if chunk:
                    # Check for frame gap if we already have data
                    if buffer and (time.time() - last_byte_time > frame_timeout):
                        break  # Previous frame complete, new frame starting
                    
                    buffer.extend(chunk)
                    last_byte_time = time.time()
                    
                    # Check if we have a complete frame
                    if len(buffer) >= 5 and self._is_complete_frame(buffer):
                        break
            else:
                # No data available, check if frame is complete
                if buffer and (time.time() - last_byte_time > frame_timeout):
                    break
                time.sleep(0.001)
        
        return bytes(buffer) if buffer else None
    
    def _is_complete_frame(self, buffer: bytes) -> bool:
        """
        Check if buffer contains a complete Modbus frame
        
        Args:
            buffer: Received data buffer
            
        Returns:
            True if frame appears complete
        """
        if len(buffer) < 5:
            return False
        
        # Exception responses are always 5 bytes
        if buffer[1] & 0x80:
            return len(buffer) >= 5
        
        # Normal responses - check function code to determine expected length
        function_code = buffer[1]
        
        if function_code in [0x03, 0x04]:  # Read registers
            if len(buffer) >= 3:
                expected_length = 5 + buffer[2]  # slave + func + count + data + 2 CRC
                return len(buffer) >= expected_length
        elif function_code in [0x01, 0x02]:  # Read coils/inputs
            if len(buffer) >= 3:
                expected_length = 5 + buffer[2]
                return len(buffer) >= expected_length
        elif function_code in [0x05, 0x06, 0x0F, 0x10]:  # Write operations
            return len(buffer) >= 8  # Fixed length responses
        
        # Default: assume complete if we have at least 5 bytes
        return len(buffer) >= 5
    
    def _parse_response(self, data: bytes) -> ModbusResponse:
        """
        Parse raw response data into ModbusResponse object
        
        Args:
            data: Raw response bytes
            
        Returns:
            Parsed ModbusResponse
        """
        if len(data) < 5:
            return ModbusResponse(timeout=True)
        
        # Extract components
        slave_id = data[0]
        function_code = data[1]
        
        # Verify CRC
        payload = data[:-2]
        received_crc = struct.unpack('<H', data[-2:])[0]  # Little endian
        calculated_crc = self.calculate_crc(payload)
        
        if received_crc != calculated_crc:
            self.stats['crc_errors'] += 1
            return ModbusResponse(
                slave_id=slave_id,
                function_code=function_code,
                crc_error=True,
                raw_frame=data
            )
        
        self.stats['responses_received'] += 1
        
        # Check for exception response
        if function_code & 0x80:
            exception_code = data[2] if len(data) > 2 else None
            self.stats['exceptions'] += 1
            return ModbusResponse(
                slave_id=slave_id,
                function_code=function_code & 0x7F,
                exception_code=exception_code,
                raw_frame=data
            )
        
        # Normal response
        response_data = data[2:-2] if len(data) > 4 else b''
        return ModbusResponse(
            slave_id=slave_id,
            function_code=function_code,
            data=response_data,
            raw_frame=data
        )
    
    # ==================== High-level Modbus Functions ====================
    
    def read_holding_registers(self, slave_id: int, start_address: int, count: int) -> Tuple[Optional[List[int]], Optional[str]]:
        """
        Read holding registers (function code 0x03)
        
        Args:
            slave_id: Slave device ID (1-247)
            start_address: Starting register address
            count: Number of registers to read
            
        Returns:
            Tuple of (register values list, error message)
        """
        return self._read_registers(slave_id, start_address, count, FunctionCode.READ_HOLDING_REGISTERS)
    
    def read_input_registers(self, slave_id: int, start_address: int, count: int) -> Tuple[Optional[List[int]], Optional[str]]:
        """
        Read input registers (function code 0x04)
        
        Args:
            slave_id: Slave device ID (1-247)
            start_address: Starting register address
            count: Number of registers to read
            
        Returns:
            Tuple of (register values list, error message)
        """
        return self._read_registers(slave_id, start_address, count, FunctionCode.READ_INPUT_REGISTERS)
    
    def _read_registers(self, slave_id: int, start_address: int, count: int, function_code: FunctionCode) -> Tuple[Optional[List[int]], Optional[str]]:
        """
        Internal method to read registers
        
        Args:
            slave_id: Slave device ID
            start_address: Starting address
            count: Number of registers
            function_code: Function code to use
            
        Returns:
            Tuple of (register values, error message)
        """
        # Build request frame
        request_data = struct.pack('>HH', start_address, count)
        frame = struct.pack('BB', slave_id, function_code.value) + request_data
        crc = self.calculate_crc(frame)
        frame += struct.pack('<H', crc)
        
        # Send request and get response
        response = self._send_frame(frame)
        
        if response.timeout:
            return None, "Request timeout"
        
        if response.crc_error:
            return None, "CRC error in response"
        
        if response.is_exception:
            return None, f"Modbus exception: {response.exception_name}"
        
        if not response.data or len(response.data) < 1:
            return None, "Invalid response data"
        
        # Parse register data
        byte_count = response.data[0]
        register_data = response.data[1:1+byte_count]
        
        if len(register_data) != byte_count:
            return None, "Incomplete register data"
        
        # Convert bytes to 16-bit registers (big endian)
        registers = []
        for i in range(0, len(register_data), 2):
            if i + 1 < len(register_data):
                reg_value = struct.unpack('>H', register_data[i:i+2])[0]
                registers.append(reg_value)
        
        return registers, None
    
    def write_single_register(self, slave_id: int, address: int, value: int) -> Tuple[bool, Optional[str]]:
        """
        Write single register (function code 0x06)
        
        Args:
            slave_id: Slave device ID (1-247)
            address: Register address
            value: Value to write (0-65535)
            
        Returns:
            Tuple of (success, error message)
        """
        # Build request frame
        request_data = struct.pack('>HH', address, value)
        frame = struct.pack('BB', slave_id, FunctionCode.WRITE_SINGLE_REGISTER.value) + request_data
        crc = self.calculate_crc(frame)
        frame += struct.pack('<H', crc)
        
        # Send request and get response
        response = self._send_frame(frame)
        
        if response.timeout:
            return False, "Request timeout"
        
        if response.crc_error:
            return False, "CRC error in response"
        
        if response.is_exception:
            return False, f"Modbus exception: {response.exception_name}"
        
        return True, None
    
    def write_multiple_registers(self, slave_id: int, start_address: int, values: List[int]) -> Tuple[bool, Optional[str]]:
        """
        Write multiple registers (function code 0x10)
        
        Args:
            slave_id: Slave device ID (1-247)
            start_address: Starting register address
            values: List of values to write
            
        Returns:
            Tuple of (success, error message)
        """
        count = len(values)
        byte_count = count * 2
        
        # Pack register values as big-endian 16-bit integers
        register_data = b''.join(struct.pack('>H', val) for val in values)
        
        # Build request frame
        request_data = struct.pack('>HHB', start_address, count, byte_count) + register_data
        frame = struct.pack('BB', slave_id, FunctionCode.WRITE_MULTIPLE_REGISTERS.value) + request_data
        crc = self.calculate_crc(frame)
        frame += struct.pack('<H', crc)
        
        # Send request and get response
        response = self._send_frame(frame)
        
        if response.timeout:
            return False, "Request timeout"
        
        if response.crc_error:
            return False, "CRC error in response"
        
        if response.is_exception:
            return False, f"Modbus exception: {response.exception_name}"
        
        return True, None
    
    # ==================== Device Testing and Discovery ====================
    
    def test_communication(self, slave_id: int) -> bool:
        """
        Test communication with a specific slave ID
        
        Args:
            slave_id: Slave ID to test
            
        Returns:
            True if communication successful
        """
        # Try reading from common register addresses
        test_addresses = [0, 1, 2048, 40001, 30001]
        
        for addr in test_addresses:
            # Try holding registers first
            registers, error = self.read_holding_registers(slave_id, addr, 1)
            if registers is not None:
                logger.debug(f"Communication test successful with slave {slave_id} at holding register {addr}")
                return True
            
            # Try input registers
            registers, error = self.read_input_registers(slave_id, addr, 1)
            if registers is not None:
                logger.debug(f"Communication test successful with slave {slave_id} at input register {addr}")
                return True
        
        return False
    
    def probe_device(self, slave_id: int, test_address: int = 65535) -> ModbusResponse:
        """
        Probe a device by sending a read request to an invalid address
        This typically generates a predictable exception response
        
        Args:
            slave_id: Slave ID to probe
            test_address: Address to test (default: 65535, likely invalid)
            
        Returns:
            ModbusResponse object
        """
        # Build request frame for impossible address
        request_data = struct.pack('>HH', test_address, 1)
        frame = struct.pack('BB', slave_id, FunctionCode.READ_HOLDING_REGISTERS.value) + request_data
        crc = self.calculate_crc(frame)
        frame += struct.pack('<H', crc)
        
        return self._send_frame(frame)
    
    def scan_slaves(self, slave_range: range, progress_callback=None) -> List[Tuple[int, ModbusResponse]]:
        """
        Scan for active Modbus slaves
        
        Args:
            slave_range: Range of slave IDs to scan
            progress_callback: Optional callback function for progress updates
            
        Returns:
            List of tuples (slave_id, response)
        """
        found_devices = []
        found_ids = set()
        
        for i, slave_id in enumerate(slave_range):
            if progress_callback:
                progress_callback(slave_id, len(slave_range))
            
            response = self.probe_device(slave_id)
            
            if response.is_valid and response.slave_id not in found_ids:
                found_ids.add(response.slave_id)
                found_devices.append((response.slave_id, response))
                logger.info(f"Found device at slave ID {response.slave_id}")
            
            time.sleep(0.01)  # Small delay between requests
        
        return found_devices
    
    # ==================== Statistics and Monitoring ====================
    
    def get_statistics(self) -> Dict[str, Any]:
        """
        Get communication statistics
        
        Returns:
            Dictionary with statistics
        """
        total_requests = self.stats['requests_sent']
        success_rate = (self.stats['responses_received'] / total_requests * 100) if total_requests > 0 else 0
        
        return {
            **self.stats,
            'success_rate': success_rate,
            'error_rate': ((self.stats['timeouts'] + self.stats['crc_errors']) / total_requests * 100) if total_requests > 0 else 0
        }
    
    def reset_statistics(self) -> None:
        """Reset all statistics"""
        self.stats = {key: 0 for key in self.stats}
    
    # ==================== Configuration Management ====================
    
    def update_config(self, **kwargs) -> None:
        """
        Update serial configuration parameters
        
        Args:
            **kwargs: Configuration parameters to update
        """
        if self.config:
            for key, value in kwargs.items():
                if hasattr(self.config, key):
                    setattr(self.config, key, value)
        else:
            # Create new config with defaults
            self.config = SerialConfig(port="", **kwargs)
    
    def get_config(self) -> Optional[SerialConfig]:
        """Get current configuration"""
        return self.config


# ==================== Utility Functions ====================

def create_modbus_master(port: str = None, baudrate: int = 9600, **kwargs) -> ModbusMaster:
    """
    Factory function to create a configured ModbusMaster instance
    
    Args:
        port: Serial port name (auto-detected if None)
        baudrate: Serial baudrate
        **kwargs: Additional serial configuration parameters
        
    Returns:
        Configured ModbusMaster instance
    """
    master = ModbusMaster()
    
    # Auto-detect port if not provided
    if port is None:
        port = master.auto_detect_port()
        if port is None:
            raise ModbusException("No suitable serial port found")
    
    # Create configuration
    config = SerialConfig(port=port, baudrate=baudrate, **kwargs)
    master.config = config
    
    return master


def scan_network(port: str = None, baudrate: int = 9600, slave_range: range = range(1, 248)) -> List[Dict[str, Any]]:
    """
    Convenience function to scan for Modbus devices on a network
    
    Args:
        port: Serial port (auto-detected if None)
        baudrate: Serial baudrate
        slave_range: Range of slave IDs to scan
        
    Returns:
        List of discovered device information
    """
    with create_modbus_master(port, baudrate) as master:
        if not master.connect():
            raise ModbusException(f"Failed to connect to {master.config.port}")
        
        devices = master.scan_slaves(slave_range)
        
        # Format results
        results = []
        for slave_id, response in devices:
            results.append({
                'slave_id': slave_id,
                'exception_name': response.exception_name,
                'timestamp': response.timestamp,
                'baudrate': baudrate
            })
        
        return results


if __name__ == "__main__":
    # Example usage and testing
    print("Modbus Master Base Class")
    print("========================")
    
    # List available ports
    ports = ModbusMaster.list_available_ports()
    print(f"Available ports: {len(ports)}")
    for port in ports:
        print(f"  {port['port']}: {port['description']}")
    
    # Find RS485 adapters
    rs485_ports = ModbusMaster.find_rs485_ports()
    print(f"Potential RS485 adapters: {len(rs485_ports)}")
    for port in rs485_ports:
        print(f"  {port['port']}: {port['description']}")