using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestParseAP.Data
{
    public class DynamicLotDetails
    {
        public string errorCode { get; set; }
        public int buyerNumber { get; set; }
        public double buyTodayBid { get; set; }
        public int currentBid { get; set; }
        public double totalAmountDue { get; set; }
        public bool sealedBid { get; set; }
        public bool firstBid { get; set; }
        public bool hasBid { get; set; }
        public bool sellerReserveMet { get; set; }
        public bool lotSold { get; set; }
        public string bidStatus { get; set; }
        public string saleStatus { get; set; }
        public string counterBidStatus { get; set; }
        public bool buyerHighBidder { get; set; }
    }

    public class LotDetails
    {
        public List<string> siteCodes { get; set; }
        public DynamicLotDetails dynamicLotDetails { get; set; }
        public string vehicleTypeCode { get; set; }
        public string lotNumberStr { get; set; }
        public bool lotSold { get; set; }
        public int ln { get; set; }
        public string mkn { get; set; }
        public string lm { get; set; }
        public int lcy { get; set; }
        public string fv { get; set; }
        public double la { get; set; }
        public double rc { get; set; }
        public string obc { get; set; }
        public double orr { get; set; }
        public string ord { get; set; }
        public string egn { get; set; }
        public string cy { get; set; }
        public string ld { get; set; }
        public string yn { get; set; }
        public string cuc { get; set; }
        public string tz { get; set; }
        public long ad { get; set; }
        public string at { get; set; }
        public int aan { get; set; }
        public double hb { get; set; }
        public int ss { get; set; }
        public string bndc { get; set; }
        public double bnp { get; set; }
        public bool sbf { get; set; }
        public string ts { get; set; }
        public string stt { get; set; }
        public string td { get; set; }
        public string tgc { get; set; }
        public string dd { get; set; }
        public string tims { get; set; }
        public List<string> lic { get; set; }
        public string gr { get; set; }
        public string dtc { get; set; }
        public string al { get; set; }
        public string adt { get; set; }
        public int ynumb { get; set; }
        public int phynumb { get; set; }
        public bool bf { get; set; }
        public int ymin { get; set; }
        public double @long { get; set; }
        public double lat { get; set; }
        public string zip { get; set; }
        public bool offFlg { get; set; }
        public string locCountry { get; set; }
        public string locCity { get; set; }
        public string locState { get; set; }
        public string tsmn { get; set; }
        public string htsmn { get; set; }
        public string tmtp { get; set; }
        public bool vfs { get; set; }
        public double myb { get; set; }
        public string lmc { get; set; }
        public string lcc { get; set; }
        public string sdd { get; set; }
        public string bstl { get; set; }
        public string lcd { get; set; }
        public string clr { get; set; }
        public string ft { get; set; }
        public string hk { get; set; }
        public string drv { get; set; }
        public string ess { get; set; }
        public bool slfg { get; set; }
        public string lsts { get; set; }
        public string snbr { get; set; }
        public bool showSeller { get; set; }
        public bool sstpflg { get; set; }
        public string std { get; set; }
        public bool isInsCpny { get; set; }
        public string vehTypDesc { get; set; }
        public string syn { get; set; }
        public bool ifs { get; set; }
        public bool ils { get; set; }
        public bool pbf { get; set; }
        public double crg { get; set; }
        public long lu { get; set; }
        public string brand { get; set; }
        public bool mof { get; set; }
    }

    public class Data
    {
        public LotDetails lotDetails { get; set; }
    }

    public class Car_atr
    {
        public int returnCode { get; set; }
        public string returnCodeDesc { get; set; }
        public Data data { get; set; }
    }
}
