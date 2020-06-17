using Prism.Mvvm;
using System;

namespace QuickOutputCert.Models
{
    public class DataModel : BindableBase
    {
        public string Index { get; set; }
        public string ProdName { get; set; }
        public string Hscode { get; set; }
        public string Evaluation { get; set; }
        public string Place { get; set; }
        public string Manufacturer { get; set; }
        public string ProdDate { get; set; }
        public string DeviceNo { get; set; }
        public string DeviceModel { get; set; }
        public string Count { get; set; }
        public string UseYear { get; set; }
        public string HowUse { get; set; }
        public string SealNo { get; set; }
        public string DeviceValue { get; set; }
        public string Check1 { get; set; }
        public string Check2 { get; set; }
        public string DeviceCheck1 { get; set; }
        public string DeviceCheck2 { get; set; }
        public string DeviceCheck3 { get; set; }
        public string DeviceCheck4 { get; set; }
        public string DeviceCheck5 { get; set; }
        public string DeviceCheck6 { get; set; }
        public string DeviceCheck7 { get; set; }
        public string DeviceCheck8 { get; set; }
        public string DeviceCheck9 { get; set; }
        public string DeviceCheck10 { get; set; }
        public string DeviceCheck11 { get; set; }
        public string DeviceCheck12 { get; set; }
        public string Remark { get; set; }

        internal object CompareTo(DataModel y)
        {
            throw new NotImplementedException();
        }
    }
}
