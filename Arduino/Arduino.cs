using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Ports;
using System.Diagnostics;
using System.Drawing;

namespace Uctrl.Arduino
{
    public class Arduino : IDisposable
    {
        private SerialPort serialPort;

        public Arduino()
        {
            serialPort = new SerialPort();
        }

        public bool OpenPort(string port)
        {
            try
            {
                serialPort.PortName = port;
                serialPort.BaudRate = 9600;
                serialPort.ReadTimeout = 1000;
                serialPort.WriteTimeout = 1000;
                serialPort.NewLine = "\n";
                serialPort.Open();

                return serialPort.IsOpen;
            }

            catch (IOException)
            {
                return false;
            }
        }

        public SerialPort Port
        {
            get { return serialPort; }
        }

        public void Dispose()
        {
            serialPort.Close();
            serialPort.Dispose();
        }

        public bool Send(string command)
        {
            if (!Port.IsOpen) return false;

            try
            {
                serialPort.WriteLine(command);

                return true;
            }

            catch (IOException)
            {
                return false;
            }
        }

        public bool SetLEDs(Color color)
        {
            return Send(color.ToArgb().ToString());
        }
    }
}
