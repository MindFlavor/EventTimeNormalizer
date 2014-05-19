using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ITPCfSQL.EventTimeNormalizer
{
    public class Normalizer
    {
        public delegate void OnePercentStepEventHandler(object sender);

        #region Members
        public TimeSpan Step { get; set; }
        #endregion

        #region Events
        public event System.ComponentModel.ProgressChangedEventHandler ProgressChanged;
        public event OnePercentStepEventHandler OnePercentStep;
        #endregion

        #region Constructors
        public Normalizer(
            TimeSpan tsStep)
        {
            this.Step = tsStep;
        }
        #endregion

        public DateValueGroup Normalize(DateTime dtStart, DateTime dtEnd, DateValueGroup dvgInput)
        {
            DateValueGroup dvgOutput = new DateValueGroup(dvgInput.Keys);

            DateTime dtCurrent = dtStart;

            long lPos = 0;

            double dLastValue = dvgInput.Values[0].Value;
            DateTime dLastDate = dvgInput.Values[0].Date;

            double dTotalSteps = ((dtEnd - dtStart).TotalMilliseconds) / Step.TotalMilliseconds;
            double dLastShownPerc = double.MinValue;
            double dSteps = 0.0D;

            int iSrcPos = 1;

            while (dtCurrent < dtEnd)
            {
                DateTime dtNext = dtCurrent.Add(Step);
                double dAccumulated = 0;

                while ((dtCurrent < dtNext))
                {
                    DateTime dtToAnalyze;
                    double dPreviousValue;

                    if (iSrcPos < dvgInput.Values.Count)
                    {
                        dtToAnalyze = dvgInput.Values[iSrcPos].Date;
                        dPreviousValue = dvgInput.Values[iSrcPos - 1].Value;
                    }
                    else
                    {
                        dtToAnalyze = DateTime.MaxValue;
                        dPreviousValue = dvgInput.Values[dvgInput.Values.Count - 1].Value;
                    }

                    if (dtToAnalyze < dtNext)
                    {
                        double deltaMSPerc = (dtToAnalyze - dtCurrent).TotalMilliseconds / Step.TotalMilliseconds;
                        dAccumulated += dPreviousValue * deltaMSPerc;

                        dtCurrent = dtToAnalyze;
                        iSrcPos++;
                    }
                    else
                    {
                        double deltaMSPerc = (dtNext - dtCurrent).TotalMilliseconds / Step.TotalMilliseconds;
                        dAccumulated += dPreviousValue * deltaMSPerc;

                        DateValuePair dvp = new DateValuePair(lPos++) { Date = dtNext.Subtract(Step) };
                        dvp.Value = dAccumulated;
                        dvgOutput.Add(dvp);

                        #region Calculate % completed
                        dSteps++;
                        double dCurPerc = (dSteps / dTotalSteps) * 100;
                        if (dCurPerc - dLastShownPerc > 1.0D)
                        {
                            OnProgress(new System.ComponentModel.ProgressChangedEventArgs((int)dCurPerc, null));
                            OnOnePercentStep();

                            dLastShownPerc = dCurPerc;
                        }
                        #endregion

                        dAccumulated = 0.0D;
                        dtCurrent = dtNext;
                    }
                }
            }

            return dvgOutput;
        }

        #region Event handlers
        public void OnProgress(System.ComponentModel.ProgressChangedEventArgs pcea)
        {
            if (ProgressChanged != null)
                ProgressChanged(this, pcea);
        }

        public void OnOnePercentStep()
        {
            if (OnePercentStep != null)
                OnePercentStep(this);
        }
        #endregion
    }
}
