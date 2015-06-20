using System;
using System.Drawing;
using System.Windows.Forms;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi.Tools;

namespace hello.PowerPointAddin
{
    public partial class SamplePane : UserControl, ITaskPane
    {
        #region Ctor

        public SamplePane()
        {
            InitializeComponent();
        }

        #endregion

        #region Properties

        private DateTime StartTime { get; set; }

        #endregion

        #region ITaskpane

        public void OnConnection(NetOffice.PowerPointApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            StartTime = DateTime.Now;
            buttonEnabled_Click(buttonEnabled, new EventArgs());
        }

        public void OnDisconnection()
        {
        }

        public void OnDockPositionChanged(MsoCTPDockPosition position)
        {
        }

        public void OnVisibleStateChanged(bool visible)
        {
        }

        #endregion

        #region UI Trigger

        private void buttonEnabled_Click(object sender, EventArgs e)
        {
            if (timerRunningTime.Enabled)
            {
                timerRunningTime.Enabled = false;
                buttonEnabled.Text = "Enable";
                buttonEnabled.ImageKey = "alarmclock_run.png";
                labelTime.ForeColor = Color.FromKnownColor(KnownColor.ControlText);
            }
            else
            {
                timerRunningTime.Enabled = true;
                buttonEnabled.Text = "Disable";
                buttonEnabled.ImageKey = "alarmclock_stop.png";
                labelTime.ForeColor = Color.White;
            }
        }

        private void buttonReset_Click(object sender, EventArgs e)
        {
            StartTime = DateTime.Now;
            labelTime.Text = "00:00:00";
        }

        private void timerRunningTime_Tick(object sender, EventArgs e)
        {
            var ts = DateTime.Now - StartTime;
            labelTime.Text = $"{ts.Hours:00}:{ts.Minutes:00}:{ts.Seconds:00}";

        }

        #endregion
    }
}
