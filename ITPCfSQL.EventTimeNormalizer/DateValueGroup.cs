using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ITPCfSQL.EventTimeNormalizer
{
    public class DateValueGroup
    {
        public DateValueGroup(params object[] Keys)
        {
            Values = new List<DateValuePair>();
            this.Keys = Keys;
        }


        public List<DateValuePair> Values { get; protected set; }

        public object[] Keys { get; protected set; }

        public DateValuePair this[int i]
        {
            get { return Values[i]; }
            set
            {
                Values[i] = value;
            }
        }

        public void Add(DateValuePair dvp)
        { 
            Values.Add(dvp);
        }

        public bool MatchKeys(object[] KeysToMatch)
        {
            if (Keys.Length != KeysToMatch.Length)
                return false;

            for (int i = 0; i < Keys.Length; i++)
            {
                if (!(Keys[i].Equals(KeysToMatch[i])))
                    return false;
            }

            return true;
        }

        public void SortDateValues()
        {
            Values.Sort((DateValuePair r1, DateValuePair r2) =>
            {
                int i = r1.Date.CompareTo(r2.Date);
                if (i != 0)
                    return i;

                return r1.OriginalPosition.CompareTo(r2.OriginalPosition);
            });

        }
    }
}
