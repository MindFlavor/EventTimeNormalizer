using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ITPCfSQL.EventTimeNormalizer
{
    public class Normalizer
    {
        public static DateValueGroup Normalize(
            DateTime dtStart, DateTime dtEnd,
            TimeSpan tsStep,
            DateValueGroup dvgInput)
        {
            DateValueGroup dvgOutput = new DateValueGroup(dvgInput.Keys);

            DateTime dtCurrent = dtStart;

            long lPos = 0;

            double dLastValue = dvgInput.Values[0].Value;
            DateTime dLastDate = dvgInput.Values[0].Date;

            int iSrcPos = 1;


            while (dtCurrent < dtEnd)
            {
                DateTime dtNext = dtCurrent.Add(tsStep);
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
                        double deltaMSPerc = (dtToAnalyze - dtCurrent).TotalMilliseconds / tsStep.TotalMilliseconds;
                        dAccumulated += dPreviousValue * deltaMSPerc;

                        dtCurrent = dtToAnalyze;
                        iSrcPos++;
                    }
                    else
                    {
                        double deltaMSPerc = (dtNext - dtCurrent).TotalMilliseconds / tsStep.TotalMilliseconds;
                        dAccumulated += dPreviousValue * deltaMSPerc;

                        DateValuePair dvp = new DateValuePair(lPos++) { Date = dtNext.Subtract(tsStep) }; // VARIABILE!
                        dvp.Value = dAccumulated;
                        dvgOutput.Add(dvp);

                        //#region Calculate % completed
                        //dSteps++;
                        //double dCurPerc = (dSteps / lTotalSteps) * 100;
                        //if (dCurPerc - dLastShownPerc > 1.0D)
                        //{
                        //    log.InfoFormat("{0:N0}% completed.", dCurPerc);
                        //    dLastShownPerc = dCurPerc;
                        //}
                        //#endregion

                        dAccumulated = 0.0D;
                        dtCurrent = dtNext;
                    }
                }
            }

            return dvgOutput;
        }
    }
}
