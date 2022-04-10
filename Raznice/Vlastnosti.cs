using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Raznice
{
    public static class Vlastnosti
    {
        //bool aEdit = false;

        public static bool allowEdit;
        public static bool exit;
        public static string logText;

        public static string logToText 
        {
            get { return logText; }
            set { value = logText; }
        }

        /*
        {
            get { }
            set { };
        }
        */ 

        public class popisStavuRaznice
        {
            public int nStatusId { get; set; }
            public string nStatusText { get; set; }
            public int nInfoId { get; set; }
            public string nInfoText { get; set; }
            public int nErrorId { get; set; }
            public string nErrorText { get; set; }
            public string stavText { get; set; }
        }

    }

    /*
    class Counter
    {
        private int threshold;
        private int total;

        public Counter(int passedThreshold)
        {
            threshold = passedThreshold;
        }

        public void Add(int x)
        {
            total += x;
            if (total >= threshold)
            {
                OnThresholdReached(EventArgs.Empty);
                
            }
        }

        protected virtual void OnThresholdReached(EventArgs e)
        {
            EventHandler handler = ThresholdReached;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler ThresholdReached;
    }   
    */
}
