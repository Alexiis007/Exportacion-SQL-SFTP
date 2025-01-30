using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exportacion
{
    internal class Timer: IDisposable
    {
        public void Tick()
        {
            if (Enabled && Environment.TickCount >= nextTick)
            {
                Callback.Invoke(this, null);
                nextTick = Environment.TickCount + Interval;
            }
        }

        private int nextTick = 0;

        public void Start()
        {
            this.Enabled = true;
            Interval = interval;
        }

        public void Stop()
        {
            this.Enabled = false;
        }

        public event EventHandler Callback;

        public bool Enabled = false;

        private int interval = 1000;

        public int Interval
        {
            get { return interval; }
            set { interval = value; nextTick = Environment.TickCount + interval; }
        }

        public void Dispose()
        {
            this.Callback = null;
            this.Stop();
        }
    }
}
