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

        public DateValuePair Current { get; protected set; }
        public DateTime CurrentStep { get; protected set; }

        protected double AccumulatedValue { get; set; }
        protected DateTime CurrentTime { get; set; }
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

        public void Start(DateValuePair Start)
        {
            this.Current = Start;
            CurrentStep = Start.Date;
            AccumulatedValue = 0.0D;
            CurrentTime = Start.Date;
        }

        public List<DateValuePair> Push(DateValuePair dvp)
        {
            List<DateValuePair> lDVPs = new List<DateValuePair>();
            DateTime NextStep = CurrentStep.Add(Step);

            while (dvp.Date >= NextStep)
            {
                double dUsedMS = Math.Min((NextStep - CurrentTime).TotalMilliseconds, 1000);
                AccumulatedValue += (dUsedMS * Current.Value) / 1000.0D;
                lDVPs.Add(new DateValuePair(-1) { Date = CurrentStep, Value = AccumulatedValue });
                AccumulatedValue = 0;

                CurrentStep = NextStep;
                CurrentTime = NextStep;
                NextStep = NextStep.Add(Step);
            }

            double dRemainining = Math.Min((dvp.Date - CurrentTime).TotalMilliseconds, 1000);
            AccumulatedValue += (dRemainining * Current.Value) / 1000.0D;
            CurrentTime = dvp.Date;

            Current = dvp;

            return lDVPs;
        }
    }
}
