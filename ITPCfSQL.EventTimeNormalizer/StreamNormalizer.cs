using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ITPCfSQL.EventTimeNormalizer
{
    public class StreamNormalizer
    {
        public delegate void OnePercentStepEventHandler(object sender);

        #region Members
        public TimeSpan Step { get; set; }

        public DateTime StartDateTime { get; set; }
        public DateTime Current { get; protected set; }

        protected double AccumulatedValue { get; set; }
        protected DateValuePair Last { get; set; }
        #endregion

        #region Events
        public event System.ComponentModel.ProgressChangedEventHandler ProgressChanged;
        public event OnePercentStepEventHandler OnePercentStep;
        #endregion

        #region Constructors
        public StreamNormalizer(
            TimeSpan tsStep)
        {
            this.Step = tsStep;
        }
        #endregion

        public void Start(DateTime dtStart)
        {
            this.StartDateTime = dtStart;
            Current = dtStart;
            AccumulatedValue = 0.0D;
            Last = new DateValuePair(-1) { Date = dtStart, Value = 0.0D };
        }

        public List<DateValuePair> Push(DateValuePair dvp)
        {
            DateTime dtNext = Current.Add(Step);
            List<DateValuePair> lDVPs = new List<DateValuePair>();

            while (dvp.Date >= dtNext)
            {
                double dUsedMS = Math.Min((dtNext - dvp.Date).TotalMilliseconds, 1000);
                AccumulatedValue += (dUsedMS * dvp.Value) / 1000.0D;
                lDVPs.Add(new DateValuePair(-1) { Date = Current, Value = AccumulatedValue });
                AccumulatedValue = 0;

                dtNext = dtNext.AddMilliseconds(dUsedMS);
            }

            AccumulatedValue += ((dvp.Date-Current).TotalMilliseconds * dvp.Value) / 1000.0D;
            Last = dvp;

            return lDVPs;
        }
    }
}
