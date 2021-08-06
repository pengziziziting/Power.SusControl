using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Power.SusControl.Models
{
    public class PurPlanRow
    {
        public PurPlanRow(params object[] array)
        {

            for ( int i =0; i < array.Length; i++)
            {
                var dtStr = Convert.ToString(array[i]);
                if ( string.IsNullOrEmpty(dtStr) )
                    continue;
                this[i+1] = DateEntity.Create(array[i]);
            }
        }
        public string Id;
        public string ProjectName;
        public string DeviceCode;
        public DateEntity Step1;
        public DateEntity Step2;
        public DateEntity Step3;
        public DateEntity Step4;
        public DateEntity Step5;
        public DateEntity Step6;
        public DateEntity Step7;
        public DateEntity Step8;
        public DateEntity this[int i]
        {
            set
            {
                if ( i < 1 && i > 8 )
                    throw new IndexOutOfRangeException();

                switch ( i )
                {
                    case 1:
                        this.Step1 = value;
                        break;
                    case 2:
                        this.Step2 = value;
                        break;
                    case 3:
                        this.Step3 = value;
                        break;
                    case 4:
                        this.Step4 = value;
                        break;
                    case 5:
                        this.Step5 = value;
                        break;
                    case 6:
                        this.Step6 = value;
                        break;
                    case 7:
                        this.Step7 = value;
                        break;
                    case 8:
                        this.Step8 = value;
                        break;
                    default:
                        throw new IndexOutOfRangeException();

                }
            }
            get
            {
                if ( i < 1 && i > 8 )
                    throw new IndexOutOfRangeException();
                DateEntity entity;
                switch ( i )
                {
                    case 1:
                        entity=this.Step1;
                        break;
                    case 2:
                        entity = this.Step2;
                        break;
                    case 3:
                        entity = this.Step3;
                        break;
                    case 4:
                        entity = this.Step4;
                        break;
                    case 5:
                        entity = this.Step5;
                        break;
                    case 6:
                        entity = this.Step6;
                        break;
                    case 7:
                        entity = this.Step7;
                        break;
                    case 8:
                        entity = this.Step8;
                        break;
                    default:
                        throw new IndexOutOfRangeException();
                }
                /*
                if ( entity == null )
                    NewLife.Log.XTrace.WriteLine("DataEntity未初始化");*/
                return entity;
            }
        }

    }
    public class DateEntity
    {
        public static DateEntity Create(object obj)
        {
            var instance = new DateEntity();
            instance.RawDate = Convert.ToDateTime(obj);
            instance.HasValue = instance.RawDate.HasValue;
            if ( instance.RawDate.HasValue )
            { 
                instance.year= instance.RawDate.Value.Year;
                instance.month = instance.RawDate.Value.Month;
                instance.day = instance.RawDate.Value.Day;
                instance.Value = new DateTime(instance.year, instance.month, instance.day);
            }
            return instance;
        }
        public bool HasValue;
        int year;
        int month;
        int day;
        DateTime? RawDate;        
        public DateTime Value;
        public int IntervalMin;
        public int IntervalAvg;
        public int IntervalMax;
        public override string ToString()
        {
            if ( RawDate.HasValue )
            {
                return string.Format("{0}-{1:D2}-{2:D2}",year,month,day);
            }
            else
                return null;
        }
    }
}
