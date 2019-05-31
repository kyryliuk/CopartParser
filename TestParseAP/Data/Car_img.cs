using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestParseAP.Data
{
    public class THUMBNAILIMAGE
    {
        public string url { get; set; }
        public string imageType { get; set; }
        public int sequenceNumber { get; set; }
        public bool swiftFlag { get; set; }
        public string imageTypeDescription { get; set; }
        public bool highRes { get; set; }
    }

    public class FULLIMAGE
    {
        public string url { get; set; }
        public string imageType { get; set; }
        public int sequenceNumber { get; set; }
        public bool swiftFlag { get; set; }
        public string imageTypeDescription { get; set; }
        public bool highRes { get; set; }
    }

    public class HIGHRESOLUTIONIMAGE
    {
        public string url { get; set; }
        public string imageType { get; set; }
        public int sequenceNumber { get; set; }
        public bool swiftFlag { get; set; }
        public string imageTypeDescription { get; set; }
        public bool highRes { get; set; }
    }

    public class ImagesList
    {
        public List<THUMBNAILIMAGE> THUMBNAIL_IMAGE { get; set; }
        public List<FULLIMAGE> FULL_IMAGE { get; set; }
        public List<HIGHRESOLUTIONIMAGE> HIGH_RESOLUTION_IMAGE { get; set; }
    }

    public class DataImg
    {
        public object lotDetails { get; set; }
        public ImagesList imagesList { get; set; }
    }

    public class Car_img
    {
        public int returnCode { get; set; }
        public string returnCodeDesc { get; set; }
        public DataImg data { get; set; }
    }
}
