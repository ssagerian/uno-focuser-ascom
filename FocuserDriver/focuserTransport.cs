using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ASCOM.ssagerianunofocuser.Focuser
{
    internal sealed class UnoFocuserTransport : IDisposable
    {
        private readonly object syncRoot = new object();
        private SerialPort serialPort;

        public string PortName { get; set; } = "COM6";
        public int BaudRate { get; set; } = 19200;
        public int ReadTimeoutMs { get; set; } = 1500;
        public int WriteTimeoutMs { get; set; } = 1500;

        public bool IsConnected => serialPort != null && serialPort.IsOpen;

        public void Connect()
        {
            lock (syncRoot)
            {
                if (IsConnected)
                    return;

                serialPort = new SerialPort(PortName, BaudRate, Parity.None, 8, StopBits.One)
                {
                    Handshake = Handshake.None,
                    NewLine = "\r\n",
                    ReadTimeout = ReadTimeoutMs,
                    WriteTimeout = WriteTimeoutMs,
                    DtrEnable = false,
                    RtsEnable = false
                };

                serialPort.Open();

                // Give the controller a moment in case USB serial toggles reset behavior.
                Thread.Sleep(2000);

                // Flush any startup noise.
                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();

                // Optional handshake check
                string id = SendCommand("ID");
                if (!id.StartsWith("OK ID ", StringComparison.OrdinalIgnoreCase))
                {
                    Disconnect();
                    throw new InvalidOperationException($"Unexpected device response to ID: '{id}'");
                }
            }
        }

        public void Disconnect()
        {
            lock (syncRoot)
            {
                if (serialPort != null)
                {
                    try
                    {
                        if (serialPort.IsOpen)
                            serialPort.Close();
                    }
                    finally
                    {
                        serialPort.Dispose();
                        serialPort = null;
                    }
                }
            }
        }

        public string SendCommand(string command)
        {
            lock (syncRoot)
            {
                EnsureConnected();

                if (serialPort == null)
                    throw new InvalidOperationException("Serial port not initialized.");

                serialPort.DiscardInBuffer();
                serialPort.Write(command + "\r\n");

                string reply = serialPort.ReadLine().Trim();

                return reply;
            }
        }

        public void SendCommandExpectOk(string command)
        {
            string reply = SendCommand(command);

            if (!reply.Equals("OK", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Command '{command}' failed. Reply: '{reply}'");
            }
        }

        public int GetPosition()
        {
            string reply = SendCommand("GP");
            return ParseIntReply(reply, "OK POS ");
        }

        public bool GetIsMoving()
        {
            string reply = SendCommand("GI");
            return ParseIntReply(reply, "OK MOV ") != 0;
        }

        public bool GetHomeState()
        {
            string reply = SendCommand("GH");
            return ParseIntReply(reply, "OK HOME ") != 0;
        }

        public bool GetEnabledState()
        {
            string reply = SendCommand("GE");
            return ParseIntReply(reply, "OK EN ") != 0;
        }

        public int GetStepSize()
        {
            string reply = SendCommand("GS");
            return ParseIntReply(reply, "OK STEP ");
        }

        public void SetStepSize(int value)
        {
            SendCommandExpectOk($"SS {value}");
        }

        public int GetMaxPosition()
        {
            string reply = SendCommand("GM");
            return ParseIntReply(reply, "OK MAX ");
        }

        public void SetMaxPosition(int value)
        {
            SendCommandExpectOk($"SM {value}");
        }

        public void MoveRelative(int delta)
        {
            SendCommandExpectOk($"MV {delta}");
        }

        public void MoveAbsolute(int position)
        {
            SendCommandExpectOk($"MA {position}");
        }

        public void Stop()
        {
            SendCommandExpectOk("ST");
        }

        public void Enable(bool enable)
        {
            SendCommandExpectOk(enable ? "EN 1" : "EN 0");
        }

        public void ZeroIfHome()
        {
            SendCommandExpectOk("HZ");
        }

        private void EnsureConnected()
        {
            if (!IsConnected)
                throw new InvalidOperationException("Transport is not connected.");
        }

        private static int ParseIntReply(string reply, string prefix)
        {
            if (!reply.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Unexpected reply: '{reply}'");

            string valueText = reply.Substring(prefix.Length).Trim();

            if (!int.TryParse(valueText, out int value))
                throw new InvalidOperationException($"Could not parse integer from reply: '{reply}'");

            return value;
        }

        public void Dispose()
        {
            Disconnect();
        }
    }
}