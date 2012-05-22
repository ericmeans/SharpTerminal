using System;
using System.Text;
using System.Threading;
using System.Collections;
using System.Runtime.InteropServices;

namespace SerialComm
{
	/// <summary>
	/// Summary description for SerialPort.
	/// </summary>
	public class SerialPort : IDisposable
	{
		private int _FileHandle = Constants.InvalidHandleValue;
		private const uint _RxBufferSize = 1024;
		private const uint _TxBufferSize = 512;

		#region Public Enums
		public enum BaudRates
		{
			CBR_110 = 110,
			CBR_300 = 300,
			CBR_600 = 600,
			CBR_1200 = 1200,
			CBR_2400 = 2400,
			CBR_4800 = 4800,
			CBR_9600 = 9600,
			CBR_14400 = 14400,
			CBR_19200 = 19200,
			CBR_38400 = 38400,
			CBR_56000 = 56000,
			CBR_57600 = 56700,
			CBR_115200 = 115200,
			CBR_128000 = 128000,
			CBR_256000 = 256000
		}

		public enum Parities : byte
		{
			None = 0,
			Odd = 1,
			Even = 2,
			Mark = 3,
			Space = 4
		}

		public enum DataSizes : byte
		{
			Four = 4,
			Five = 5,
			Six = 6,
			Seven = 7,
			Eight = 8
		}

		public enum StopSizes : byte
		{
			One = 0,
			OnePointFive = 1,
			Two = 2
		}

		public enum FlowControlTypes : byte
		{
			None,
			Hardware,
			XonXoff
		}
		#endregion

		#region Public Fields
		public string ComPort = "";
		public BaudRates BaudRate = 0;
		public Parities Parity = 0;
		public DataSizes DataBits = 0;
		public StopSizes StopBits = 0;
		public FlowControlTypes FlowControl = FlowControlTypes.None;
		#endregion

		#region Events
		public event StatusUpdateEventHandler StatusUpdate;
		public event ReceiveEventHandler Receive;
		public event WriteEventHandler WriteComplete;
		public event ErrorEventHandler Error;
		#endregion

		#region Private Fields
		private System.Threading.Thread _ReadOperationsThread = null;
		private System.Threading.Thread _WriteOperationsThread = null;
		private Queue _WriteQueue = Queue.Synchronized(new Queue());
		private ManualResetEvent _WriteHandle;
		private volatile bool _RunThreads = true;
		#endregion

		#region Properties
		private bool _Connected = false;
		public bool Connected 
		{
			get 
			{
				return _Connected;
			}
		}

		private bool _RtsState = true;
		public bool RtsState
		{
			get
			{
				return _RtsState;
			}
			set
			{
				// Only valid if the handshaking is disabled
				if (this.FlowControl != FlowControlTypes.Hardware)
				{
					if (value != _RtsState)
					{
						// Set RTS high
						uint newState = (uint)(value ? Constants.EscapeCommFunctions.SetRts : Constants.EscapeCommFunctions.ClearRts);
						if (EscapeCommFunction(_FileHandle, newState))
						{
							_RtsState = value;
						}
						else 
						{
							int lastError = GetLastError();
							HandleError(string.Format("Failed to set RTS state; EscapeCommFunction returned error code {0}: {1}", lastError, GetErrorText(lastError)));											
						}
					}
				}
				else
				{
					throw new InvalidOperationException("You cannot change the line state when hardware flow control is enabled.");
				}
			}
		}
		#endregion

		#region Constructors/Destructors/Dispose
		public SerialPort()
		{
		}

		public SerialPort(string port, BaudRates baudRate, Parities parity, DataSizes dataBits, StopSizes stopBits)
		{
			this.ComPort = port;
			this.BaudRate = baudRate;
			this.Parity = parity;
			this.DataBits = dataBits;
			this.StopBits = stopBits;
		}

		~SerialPort()
		{
			this.Dispose();
		}

		#region IDisposable Members
		public void Dispose()
		{
			if (_Connected)
			{
				this.Disconnect();
			}

			if (_FileHandle != Constants.InvalidHandleValue)
			{
				CloseHandle(_FileHandle);
			}

			GC.SuppressFinalize(this);
		}
		#endregion
		#endregion

		public void Connect()
		{
			if (_Connected)
			{
				throw new InvalidOperationException("You must disconnect from the current COM port before connecting to another one.");
			}

			if (DataBits == DataSizes.Five && StopBits == StopSizes.Two)
			{
				throw new InvalidOperationException("5 data bits and 2 stop bits is an invalid combination.");
			}
			if ((DataBits == DataSizes.Six || DataBits == DataSizes.Seven || DataBits == DataSizes.Eight) && StopBits == StopSizes.OnePointFive)
			{
				throw new InvalidOperationException("6, 7, or 8 data bits and 1.5 stop bits is an invalid combination.");
			}

			_RunThreads = false;

			// Try just opening the COM port.
			string ComPort = this.ComPort;
			_FileHandle = CreateFile(ComPort, 
				Constants.GenericRead | Constants.GenericWrite,
				0,		// No sharing
				0,		// No security attributes
				Constants.OpenExisting,
				Constants.FileFlagOverlapped,
				0);	// No template file
			if (_FileHandle == Constants.InvalidHandleValue || _FileHandle == 0)
			{
				// That didn't work, try to open \\.\COMXX instead of just COMXX
				ComPort = (this.ComPort.StartsWith(@"\\.\") ? this.ComPort : @"\\.\" + this.ComPort);
				_FileHandle = CreateFile(ComPort, 
					Constants.GenericRead | Constants.GenericWrite,
					0,		// No sharing
					0,		// No security attributes
					Constants.OpenExisting,
					Constants.FileFlagOverlapped,
					0);	// No template file
				if (_FileHandle == Constants.InvalidHandleValue || _FileHandle == 0)
				{
					int lastError = GetLastError();
					if (lastError == 2)
					{
						throw new Exception("Specified COM port does not exist.");
					}
					else
					{
						throw new Exception(string.Format("CreateFile returned error code {0}: {1}", lastError, GetErrorText(lastError)));
					}
				}
			}

			// Clear any errors on the channel
			uint error = 0;
			CommStatus s;
			ClearCommError(_FileHandle, out error, out s);

			// Clear the buffers
			PurgeComm(_FileHandle, (uint)Constants.Purges.ClearTransmitBuffer | (uint)Constants.Purges.ClearReceiveBuffer);

			// Get the current settings
			DCB deviceConfigBlock = new DCB();
			GetCommState(_FileHandle, out deviceConfigBlock);

			// Update the settings
			string state = string.Format("baud={0} parity={1} data={2} stop={3}", 
				BaudRate, "NOEMS".Substring((int)Parity, 1), DataBits, StopBits);
			BuildCommDCB(state, ref deviceConfigBlock);
			switch (this.FlowControl)
			{
				case FlowControlTypes.None:
					deviceConfigBlock.RtsControl = (uint)RtsControlTypes.Enable;
					deviceConfigBlock.DtrControl = (uint)DtrControlTypes.Enable;
					deviceConfigBlock.InX = 0;
					deviceConfigBlock.OutX = 0;
					break;
				case FlowControlTypes.Hardware:
					deviceConfigBlock.RtsControl = (uint)RtsControlTypes.Handshake;
					deviceConfigBlock.DtrControl = (uint)DtrControlTypes.Enable;
					deviceConfigBlock.InX = 0;
					deviceConfigBlock.OutX = 0;
					break;
				case FlowControlTypes.XonXoff:
					deviceConfigBlock.RtsControl = (uint)RtsControlTypes.Enable;
					deviceConfigBlock.DtrControl = (uint)DtrControlTypes.Enable;
					deviceConfigBlock.InX = 1;
					deviceConfigBlock.OutX = 1;
					break;
			}
			error = SetCommState(_FileHandle, ref deviceConfigBlock);

			if (error == 0)
			{
				throw new Exception(string.Format("Unable to set COM port state: {0}", GetErrorText(GetLastError())));
			}

			// Set up the Tx/Rx buffers
			SetupComm(_FileHandle, _RxBufferSize, _TxBufferSize);

			// Set the timeouts
			CommTimeouts timeouts;
			timeouts.ReadIntervalTimeout = 0;
			timeouts.ReadTotalTimeoutMultiplier = 0;
			timeouts.ReadTotalTimeoutConstant = 70;
			timeouts.WriteTotalTimeoutMultiplier = 10;
			timeouts.WriteTotalTimeoutConstant = 100;
			SetCommTimeouts	(_FileHandle, ref timeouts);

			_RunThreads = true;

			// Set up the read ops thread
			if (_ReadOperationsThread != null)
			{
				_ReadOperationsThread.Abort();
			}
			_ReadOperationsThread = new System.Threading.Thread(new System.Threading.ThreadStart(ReadMain));
			_ReadOperationsThread.IsBackground = true;
			_ReadOperationsThread.Name = "SerialComm Read Operations Thread";
			_ReadOperationsThread.Start();

			// Set up the write ops thread
			if (_WriteOperationsThread != null)
			{
				_WriteOperationsThread.Abort();
			}
			_WriteHandle = new ManualResetEvent(false);
			_WriteOperationsThread = new System.Threading.Thread(new System.Threading.ThreadStart(WriteMain));
			_WriteOperationsThread.IsBackground = true;
			_WriteOperationsThread.Name = "SerialComm Write Operations Thread";
			_WriteOperationsThread.Start();

			_Connected = true;
		}

		public void Disconnect()
		{
			_RunThreads = false;
			if (!_WriteOperationsThread.Join(250))
			{
				// Didn't quit in time, terminate it
				_WriteOperationsThread.Abort();
			}
			if (!_ReadOperationsThread.Join(250))
			{
				// Didn't quit in time, terminate it
				_ReadOperationsThread.Abort();
			}

			// Close our file handle
			CloseHandle(_FileHandle);
			_FileHandle = Constants.InvalidHandleValue;

			_Connected = false;
		}

		public void Write(byte[] buffer)
		{
			_WriteQueue.Enqueue(buffer);
			// Signal the write thread
			_WriteHandle.Set();
		}

		private void WriteMain()
		{
			Overlapped writeOverlapped = new Overlapped();
			writeOverlapped.Event = Constants.InvalidHandleValue;
			System.Threading.WaitHandle waitHandle = new System.Threading.ManualResetEvent(false);
			try
			{
				writeOverlapped.Event = CreateEvent(0, true, false, null);
				waitHandle.Handle = new IntPtr(writeOverlapped.Event);

				while (_RunThreads)
				{
					// Wait on something to be added to the queue
					if (_WriteHandle.WaitOne())
					{
						while (_WriteQueue.Count > 0)
						{
							byte[] buffer = (byte[])_WriteQueue.Dequeue();
							uint bytesWritten;
							// Issue write.
							if (!WriteFile(_FileHandle, buffer, Convert.ToUInt32(buffer.Length), out bytesWritten, ref writeOverlapped)) 
							{
								if (GetLastError() != Constants.ErrorIOPending) 
								{ 
									// WriteFile failed, but isn't delayed. Report error and abort.
									HandleError(string.Format("Write operation for data '{0}' failed.", Encoding.UTF8.GetString(buffer)));
								}
								else
								{
									// Write is pending.
									if (waitHandle.WaitOne(100, false))
									{
										if (!GetOverlappedResult(_FileHandle, ref writeOverlapped, out bytesWritten, false))
										{
											int lastError = GetLastError();
											HandleError(string.Format("GetOverlappedResult for write operation returned error code {0}: {1}", lastError, GetErrorText(lastError)));											
										}
										else
										{
											// Write operation completed successfully.
											if (WriteComplete != null)
											{
												WriteComplete(this, new WriteEventArgs(buffer));
											}
										}
									}
									else
									{
										HandleError("Write handle never received a signal.");
									}
								}
							}
							else
							{
								// WriteFile completed immediately.
								if (WriteComplete != null)
								{
									WriteComplete(this, new WriteEventArgs(buffer));
								}
							}				
						}	// Queue
						_WriteHandle.Reset();
					} // Wait Handle
				}	// RunThreads
			}
			finally
			{
				// Clean up the OS handles
				waitHandle.Close();
				_WriteHandle.Close();
				CloseHandle(writeOverlapped.Event);
			}
		}

		private void ReadMain()
		{
			Overlapped readOverlapped = new Overlapped();
			readOverlapped.Event = Constants.InvalidHandleValue;
			Overlapped commOverlapped = new Overlapped();
			commOverlapped.Event = Constants.InvalidHandleValue;
			System.Threading.WaitHandle[] waitHandles = new System.Threading.WaitHandle[2];
			try
			{
				// Set up the comm event mask
				uint dwStoredFlags = (uint)Constants.SerialCommEvents.BreakReceived | 
														 (uint)Constants.SerialCommEvents.CtsStateChanged |
														 (uint)Constants.SerialCommEvents.DsrStateChanged |
														 (uint)Constants.SerialCommEvents.LineStatusError |
														 (uint)Constants.SerialCommEvents.RingDetected |
														 (uint)Constants.SerialCommEvents.RlsdStateChanged;
				if (!SetCommMask(_FileHandle, dwStoredFlags))
				{
					HandleError("There was an error when setting the Comm Mask.");
					return;
				}

				readOverlapped.Event = CreateEvent(0, true, false, null);
				commOverlapped.Event = CreateEvent(0, true, false, null);

				// Set up the wait events
				waitHandles[0] = new System.Threading.ManualResetEvent(false);
				waitHandles[0].Handle = new IntPtr(readOverlapped.Event);
				waitHandles[1] = new System.Threading.ManualResetEvent(false);
				waitHandles[1].Handle = new IntPtr(commOverlapped.Event);

				bool waitingOnRead = false;
				byte[] readBuffer = new byte[_RxBufferSize];
				uint bytesRead;

				bool waitingOnStatus = false;
				uint commStatus = 0;
				uint commBytesRead = 0;

				while (_RunThreads)
				{
					// Issue a read operation
					if (!waitingOnRead)
					{
						if (!ReadFile(_FileHandle, readBuffer, _RxBufferSize, out bytesRead, ref readOverlapped)) 
						{
							int lastError = GetLastError();
							if (lastError == Constants.ErrorIOPending)
							{
								// Read delayed - no data yet
								waitingOnRead = true;
							}
							else
							{
								HandleError(string.Format("ReadFile returned error code {0}: {1}", lastError, GetErrorText(lastError)));
							}
						}
						else 
						{
							// read completed immediately
							ReportInput(readBuffer, bytesRead);
						}
					}

					// Issue a comm status update
					if (waitingOnStatus) 
					{
						if (!WaitCommEvent(_FileHandle, out commStatus, ref commOverlapped)) 
						{
							int lastError = GetLastError();
							if (lastError == Constants.ErrorIOPending)
							{
								waitingOnStatus = true;
							}
							else
							{
								HandleError(string.Format("WaitCommEvent returned error code {0}: {1}", lastError, GetErrorText(lastError)));
							}
						}
						else
						{
							// WaitCommEvent returned immediately.
							// Deal with status event as appropriate.
							ReportStatusEvent(commStatus);
						}
					}

					// Wait on the handles to be signaled
					int result = System.Threading.WaitHandle.WaitAny(waitHandles, 100, false);
					if (result == System.Threading.WaitHandle.WaitTimeout)
					{
						// Wait timed out, loop around
					}
					else if (result == 0)
					{
						// Read event signalled
						if (!GetOverlappedResult(_FileHandle, ref readOverlapped, out bytesRead, false))
						{
							int lastError = GetLastError();
							HandleError(string.Format("GetOverlappedResult for read operation returned error code {0}: {1}", lastError, GetErrorText(lastError)));
						}
						else if (bytesRead > 0)
						{
							// Read completed successfully.
							ReportInput(readBuffer, bytesRead);
						}
						// Reset flag so that another opertion can be issued.
						waitingOnRead = false;
					}
					else if (result == 1)
					{
						// Comm event signalled
						if (!GetOverlappedResult(_FileHandle, ref commOverlapped, out commBytesRead, false))
						{
							int lastError = GetLastError();
							HandleError(string.Format("GetOverlappedResult for Comm Status returned error code {0}: {1}", lastError, GetErrorText(lastError)));
						}
						else
						{
							// Status event is stored in the event flag
							// specified in the original WaitCommEvent call.
							// Deal with the status event as appropriate.
							ReportStatusEvent(commStatus);
						}
						// Reset flag so that another opertion can be issued.
						waitingOnStatus = false;
					}
				}
			}
			finally
			{
				// Clean up the OS handles
				waitHandles[0].Close();
				waitHandles[1].Close();
				CloseHandle(commOverlapped.Event);
				CloseHandle(readOverlapped.Event);
			}
		}

		private void ReportInput(byte[] buffer, uint length)
		{
			if (Receive != null)
			{
				Receive(this, new ReceiveEventArgs(Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(buffer, 0, Convert.ToInt32(length)))));
			}
		}

		private void ReportStatusEvent(uint commStatus)
		{
			// Get and clear current errors on the port.
			uint errors;
			CommStatus status;
			if (!ClearCommError(_FileHandle, out errors, out status))
			{
				HandleError("Unexpected error in ClearCommError.");
				return;
			}

			// Get comm status info
			uint modemStatus;
			if (!GetCommModemStatus(_FileHandle, out modemStatus))
			{
				HandleError("Unexpected error in GetCommModemStatus.");
				return;
			}

			// Get error flags.
			uint errBreak = errors & (uint)Constants.SerialCommErrors.BreakDetected;
			uint errFrame = errors & (uint)Constants.SerialCommErrors.ReceiveFramingError;
			uint errRxOverrun = errors & (uint)Constants.SerialCommErrors.ReceiveQueueOverflow;
			uint errTxFull = errors & (uint)Constants.SerialCommErrors.TransmitQueueFull;
			uint errOverrun = errors & (uint)Constants.SerialCommErrors.ReceiveOverrun;
			uint errParity = errors & (uint)Constants.SerialCommErrors.ReceiveParityError;

			uint CtsStatus = modemStatus & (uint)Constants.ModemStatus.CtsOn;
			uint DsrStatus = modemStatus & (uint)Constants.ModemStatus.DsrOn;
			uint RingStatus = modemStatus & (uint)Constants.ModemStatus.RingOn;
			uint RlsdStatus = modemStatus & (uint)Constants.ModemStatus.RlsdOn;

			// Update any clients with the new line status
			if (StatusUpdate != null)
			{
				StatusUpdate(this, new StatusUpdateEventArgs(CtsStatus != 0,
																										 DsrStatus != 0,
																										 RingStatus != 0,
																										 RlsdStatus != 0,
																										 errBreak != 0,
																										 errFrame != 0,
																										 errRxOverrun != 0,
																										 errTxFull != 0,
																										 errOverrun != 0,
																										 errParity != 0));
			}
		}

		private void HandleError(string message)
		{
			HandleError(message, null);
		}

		private void HandleError(string message, Exception ex)
		{
			if (Error != null)
			{
				Error(this, new ErrorEventArgs(message, ex));
			}
		}

		private string GetErrorText(int errorCode)
		{
			StringBuilder rc = new StringBuilder(256);
			int ret = FormatMessage(0x1000, 0, errorCode, 0, rc, 256, 0);
			
			if (ret > 0)
			{
				return rc.ToString();
			}
			else
			{
				return string.Format("Error code {0} not found.", errorCode);
			}
		}

		#region Win32 functions and constants
		[DllImport("kernel32.dll", SetLastError=true)]
		private static extern int CreateFile(string lpFileName, uint dwDesiredAccess,
			uint dwShareMode, int lpSecurityAttributes, uint dwCreationDisposition,
			uint dwFlagsAndAttributes, int hTemplateFile);

		[DllImport("kernel32.dll")]
		private static extern bool SetCommMask(int hFile, uint dwEvtMask);

		[DllImport("kernel32.dll")]
		private static extern bool ClearCommError(int hFile, out uint lpErrors,
			out CommStatus lpStat);

		[DllImport("kernel32.dll")]
		private static extern bool PurgeComm(int hFile, uint dwFlags);

		[DllImport("kernel32.dll")]
		private static extern bool GetCommState(int hFile, out DCB lpDCB);

		[DllImport("kernel32.dll")]
		private static extern bool BuildCommDCB(string lpDef, ref DCB lpDCB);

		[DllImport("kernel32.dll")]
		private static extern uint SetCommState(int hFile, [In] ref DCB lpDCB);

		[DllImport("kernel32.dll")]
		static extern bool EscapeCommFunction(int hFile, uint dwFunc);

		[DllImport("kernel32.dll")]
		private static extern bool SetupComm(int hFile, uint dwInQueue, uint dwOutQueue);

		[DllImport("kernel32.dll")]
		private static extern int SetCommTimeouts(int hFile, ref CommTimeouts lpCommTimeouts);
																						
		[DllImport("kernel32.dll")]
		private static extern bool ReadFile(int hFile, byte[] lpBuffer,
			uint nNumberOfBytesToRead, out uint lpNumberOfBytesRead, ref Overlapped lpOverlapped);

		[DllImport("kernel32.dll")]
		static extern bool WriteFile(int hFile, byte[] lpBuffer,
			uint nNumberOfBytesToWrite, out uint lpNumberOfBytesWritten,
			ref Overlapped lpOverlapped);

		[DllImport("kernel32.dll")]
		private static extern int CreateEvent(int lpEventAttributes, bool bManualReset,
			bool bInitialState, string lpName);
											
		[DllImport("kernel32.dll", SetLastError=true)]
		private static extern int CloseHandle(int hObject);
			
		[DllImport("kernel32.dll")]
		private static extern bool GetCommModemStatus(int hFile, out uint lpModemStat);
												
		[DllImport("kernel32.dll")]
		static extern bool WaitCommEvent(int hFile, out uint lpEvtMask,
			ref Overlapped lpOverlapped);

		[DllImport("kernel32.dll")]
		private static extern bool GetOverlappedResult(int hFile, ref Overlapped
			lpOverlapped, out uint lpNumberOfBytesTransferred, bool bWait);

		[DllImport("kernel32.dll")]
		private static extern int FormatMessage(int dwFlags, int lpSource, int dwMessageId,
			int dwLanguageId, StringBuilder lpBuffer, int nSize, int Arguments);

		[DllImport("kernel32.dll")]
		private static extern int GetLastError();
																																				 
		private abstract class Constants
		{
			public const uint GenericRead = 0x80000000;
			public const uint GenericWrite = 0x40000000;
			public const int OpenExisting = 3;
			public const uint FileFlagOverlapped = 0x40000000;
			public const int InvalidHandleValue = -1;

			public const int IOBufferSize = 1024;

			public const int ErrorIOIncomplete = 996;
			public const int ErrorIOPending = 997;

			[Flags()]
			public enum SerialCommEvents : uint
			{
				CharacterReceived = 0x0001,
				SpecificCharacterReceived = 0x0002,
				TransmitQueueEmpty = 0x0004,
				CtsStateChanged = 0x0008,
				DsrStateChanged = 0x0010,
				RlsdStateChanged = 0x0020,
				BreakReceived = 0x0040,
				LineStatusError = 0x0080,
				RingDetected = 0x0100,
				PrinterErrorDetected = 0x0200,
				ReceiveBuffer80PercentFull = 0x0400,
				ProviderSpecificEvent1 = 0x0800,
				ProviderSpecificEvent2 = 0x1000
			}

			[Flags()]
			public enum SerialCommErrors : uint	
			{
				ReceiveQueueOverflow = 0x0001,
				ReceiveOverrun = 0x0002,
				ReceiveParityError = 0x0004,
				ReceiveFramingError = 0x0008,
				BreakDetected = 0x0010,
				TransmitQueueFull = 0x0100
			}

			[Flags()]
			public enum Purges : uint
			{
				AbortTransmit = 0x0001,
				AbortReceive = 0x0002,
				ClearTransmitBuffer = 0x0004,
				ClearReceiveBuffer = 0x0008
			}
			
			[Flags()]
			public enum ModemStatus : uint
			{
				CtsOn = 0x0010,
				DsrOn = 0x0020,
				RingOn = 0x0040,
				RlsdOn = 0x0080
			}

			public enum EscapeCommFunctions : uint
			{
				SetXoff = 1,
				SetXon = 2,
				SetRts = 3,
				ClearRts = 4,
				SetDtr = 5,
				ClearDtr = 6,
				ResetDevice = 7,
				SetBreak = 8,
				ClearBreak = 9
			}
		}

		[Flags()]
		private enum DcbFlags : int
		{
			Binary = 0x0001,				// Enable binary mode -- always used
			Parity = 0x0002,				// Enable parity checking
			OutXCtsFlow = 0x0004,		// Monitor CTS signal and use for flow control (if used, CTS must be high to send)
			OutXDsrFlow = 0x0008,		// Monitor DSR signal and use for flow control (if used, DSR must be high to send)
			DtrControl1 = 0x0010,		// Specifies the DTR flow control.  Three possible values: DISABLE (DTR always low), 
															//  ENABLE (DTR always high), HANDSHAKE (DTR automatically set/cleared for handshaking)
			DtrControl2 = 0x0020,
			DsrSensitivity = 0x0040,// If used, the driver ignores any bytes received while DSR is not high.
			TransmitContinueOnXOff = 0x0080,	// Specifies whether transmission stops when the input buffer is full and the 
																				//  driver has transmitted the XoffChar character. If this member is TRUE, 
																				//  transmission continues after the input buffer has come within XoffLim 
																				//  bytes of being full and the driver has transmitted the XoffChar character 
																				//  to stop receiving bytes. If this member is FALSE, transmission does not 
																				//  continue until the input buffer is within XonLim bytes of being empty and 
																				//  the driver has transmitted the XonChar character to resume reception. 
			UseXonXoffOut = 0x0100,	// If used, transmission starts when XonChar is received and stops when XoffChar is received.
			UseXonXoffIn = 0x0200,	// If used, XoffChar is sent when the input buffer is within XoffLim bytes of being full,
															//  and XonChar is sent when the input buffer is within XonLim bytes of being empty.
			ErrorChar = 0x0400,			// Replace parity-error bytes with the ErrorChar byte.
			Null = 0x0800,					// Discard null bytes.
			RtsControl1 = 0x1000,		// Which RTS scheme to use.  May be one of the following: DISABLE (RTS always low),
															//  ENABLE (RTS always high), HANDSHAKE (automatic RTS handshaking), or 
															//  TOGGLE (RTS high whenever bytes are ready to be sent, low otherwise).
			RtsControl2 = 0x2000,
			AbortOnError = 0x4000,	// Terminate read/write ops if an error occurs.  If used, any error must be cleared
															//  by using the ClearCommError function before further reading/writing.
		}

//		[StructLayout(LayoutKind.Sequential)]
//		private struct DCB
//		{
//			public int DcbLength;	// Specifies the length of the struct, in bytes
//			public int BaudRate;	// Specifies either an actual baud rate, or an index from the BaudRates enum.
//			public DcbFlags Flags;
//			public short Reserved;
//			public short XonLimit;	// min # of bytes accepted in input buffer before XON is sent.
//			public short XoffLimit;	// max # of bytes accepted in input buffer before XOFF is sent.
//															//  The max # of bytes accepted = buffer size - XoffLimit
//			public byte ByteSize;		// The number of bits per byte
//			public byte Parity;			// Parity scheme to use
//			public byte StopBits;		// # of stop bits per byte: 1, 2, or 1.5
//			public byte XonChar;		// What character to use as the XON signal char
//			public byte XoffChar;		// What character to use as the XOFF signal char
//			public byte ErrorChar;	// What character to use to replace bytes which failed the parity check
//			public byte EofChar;		// The "End of File" character
//			public byte EvtChar;		// The character used to signal events
//			public short Reserved2;
//		}
		[StructLayout(LayoutKind.Sequential)]
		private struct DCB
		{
			public uint DcbLength;				// Length in bytes of struct
			public uint BaudRate;					// Baud rate to use for comm
			public uint Binary;						// Binary mode (skip EOF check)
			public uint EnableParity;			// Enable parity checking
			public uint OutxCtsFlow;			// Enable CTS handshaking on output
			public uint OutxDsrFlow;			// Enable DSR handshaking on output
			public uint DtrControl;				// DTR Flow Control
			public uint DsrSensitivity;		// DSR Sensitivity
			public uint TxContinueOnXoff;	// Continue transmitting after sending Xoff
			public uint OutX;							// Enable output XOn/XOff control
			public uint InX;							// Enable input XOn/XOff control
			public uint EnableErrorChar;	// Enable error replacement
			public uint StripNull;				// Enable null byte stripping
			public uint RtsControl;				// RTS Flow Control
			public uint AbortOnError;			// Abort all pending reads and writes on error
			private uint Dummy2;					// Reserved
			private short Reserved;				// Reserved
			public short XonLimit;				// Transmit XON threshold
			public short XoffLimit;				// Transmit XOFF threshold
			public byte ByteSize;					// Number of bits/byte, 4-8
			public byte Parity;						// 0-4 = None, Odd, Even, Mark, Space
			public byte StopBits;					// 0,1,2 = 1, 1.5, 2
			public char XonChar;					// Character to use for Tx/Rx XON
			public char XoffChar;					// Character to use for Tx/Rx XOFF
			public char ErrorChar;				// Character to use for replacing bytes that fail parity check
			public char EofChar;					// End of Input char
			public char EvtChar;					// Received Event char
			private short Reserved1;			// Reserved
		}

		private enum DtrControlTypes : uint
		{
			Disable = 0,
			Enable = 1,
			Handshake = 2
		}

		private enum RtsControlTypes : uint
		{
			Disable = 0,
			Enable = 1,
			Handshake = 2,
			Toggle = 3
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct CommTimeouts
		{
			public int ReadIntervalTimeout;		// Max allowed time between the arrival of any
																				//  two chars during the same read operation
			public int ReadTotalTimeoutMultiplier;
			public int ReadTotalTimeoutConstant;	// Constant used to calc total time-out
			public int WriteTotalTimeoutMultiplier;
			public int WriteTotalTimeoutConstant;
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct Overlapped
		{
			public int Internal;	// Reserved for OS use
			public int InternalHigh;	// Reserved for OS use
			public int Offset;				// Always 0 for communication devices
			public int OffsetHigh;		// Likewise
			public int Event;					// The handle to an OS event to signal when an operation is completed.
		}

		[Flags()]
		private enum CommStatusFlags : int
		{
			CtsHold = 1,	// Waiting for CTS to be asserted
			DsrHold = 2,	// Waiting for DSR to be asserted
			RlsdHold = 4,	// Waiting for RLSD to be asserted
			XoffHold = 8,	// Received an XOFF, waiting for an XON
			XoffSent = 16,// Sent an XOFF; waiting to xmit because the other system will take next char as XON.
			Eof = 32,			// Have received an end-of-file char
			Txim = 64 		// The TransmitCommChar function has queued a character for transmission at the head of the queue.
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct CommStatus
		{
			public CommStatusFlags Flags;
			public int InQueueLength;
			public int OutQueueLength;
		}
		#endregion
	} // SerialPort class

	#region Delegates and Event classes
	public delegate void StatusUpdateEventHandler(Object sender, StatusUpdateEventArgs e);
	public delegate void ReceiveEventHandler(Object sender, ReceiveEventArgs e);
	public delegate void WriteEventHandler(Object sender, WriteEventArgs e);
	public delegate void ErrorEventHandler(Object sender, ErrorEventArgs e);
	public class StatusUpdateEventArgs : EventArgs
	{
		public readonly bool CTS;
		public readonly bool DSR;
		public readonly bool Ring;
		public readonly bool CarrierDetect;
		public readonly bool BreakDetected;
		public readonly bool FramingError;
		public readonly bool RxOverflow;
		public readonly bool TxFull;
		public readonly bool Overrun;
		public readonly bool ParityError;
		public StatusUpdateEventArgs(bool cts, bool dsr, bool ring, bool rlsd, bool breakDetect, bool frame, bool rxOverflow,
																 bool txFull, bool overrun, bool parity)
		{
			this.CTS = cts;
			this.DSR = dsr;
			this.Ring = ring;
			this.CarrierDetect = rlsd;
			this.BreakDetected = breakDetect;
			this.FramingError = frame;
			this.RxOverflow = rxOverflow;
			this.TxFull = txFull;
			this.Overrun = overrun;
			this.ParityError = parity;
		}
	}

	public class ReceiveEventArgs : EventArgs
	{
		public readonly byte[] Data;
		public ReceiveEventArgs(byte[] data)
		{
			this.Data = data;
		}
	}

	public class WriteEventArgs : EventArgs
	{
		public readonly byte[] Data;
		public WriteEventArgs(byte[] data)
		{
			this.Data = data;
		}
	}

	public class ErrorEventArgs : EventArgs
	{
		public readonly Exception Exception;
		public readonly string Message;
		public ErrorEventArgs(string message)
		{
			this.Message = message;
			this.Exception = null;
		}
		public ErrorEventArgs(string message, Exception ex)
		{
			this.Message = message;
			this.Exception = ex;
		}
	}
	#endregion

} // namespace
