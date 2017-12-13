using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;

namespace HooverUnlimited.DotNetRtfWriter
{
    /// <summary>
    ///     Summary description for RtfImage
    /// </summary>
    public class RtfImage : RtfBlock
    {
        private Align _alignment;
        private string _blockHead;
        private string _blockTail;
        private float _height;
        private readonly byte[] _imgByte;
        private string _imgFname;
        private readonly ImageFileType _imgType;
        private readonly Margins _margins;
        private bool _startNewPage;
        private float _width;

        internal RtfImage(string fileName, ImageFileType type)
        {
            _imgFname = fileName;
            _imgType = type;
            _alignment = Align.None;
            _margins = new Margins();
            KeepAspectRatio = true;
            _blockHead = @"{\pard";
            _blockTail = @"}";
            _startNewPage = false;
            StartNewPara = false;

            var image = Image.FromFile(fileName);
            _width = image.Width / image.HorizontalResolution * 72;
            _height = image.Height / image.VerticalResolution * 72;

            using (var mStream = new MemoryStream())
            {
                image.Save(mStream, image.RawFormat);
                _imgByte = mStream.ToArray();
            }
        }


        internal RtfImage(MemoryStream imageStream)
        {
            _alignment = Align.Left;
            _margins = new Margins();
            KeepAspectRatio = true;
            _blockHead = @"{\pard";
            _blockTail = @"}";
            _startNewPage = false;
            StartNewPara = false;

            _imgByte = imageStream.ToArray();

            var image = Image.FromStream(imageStream);
            _width = image.Width / image.HorizontalResolution * 72;
            _height = image.Height / image.VerticalResolution * 72;

            if (image.RawFormat.Equals(ImageFormat.Png)) _imgType = ImageFileType.Png;
            else if (image.RawFormat.Equals(ImageFormat.Jpeg)) _imgType = ImageFileType.Jpg;
            else if (image.RawFormat.Equals(ImageFormat.Gif)) _imgType = ImageFileType.Gif;
            else throw new Exception("Image format is not supported: " + image.RawFormat);
        }


        public override Align Alignment
        {
            get { return _alignment; }
            set { _alignment = value; }
        }

        public override Margins Margins => _margins;

        public override bool StartNewPage
        {
            get { return _startNewPage; }
            set { _startNewPage = value; }
        }

        public bool StartNewPara { get; set; }

        public float Width
        {
            get { return _width; }
            set
            {
                if (KeepAspectRatio && _width > 0)
                {
                    var ratio = _height / _width;
                    _height = value * ratio;
                }
                _width = value;
            }
        }

        public float Heigth
        {
            get { return _height; }
            set
            {
                if (KeepAspectRatio && _height > 0)
                {
                    var ratio = _width / _height;
                    _width = value * ratio;
                }
                _height = value;
            }
        }

        public bool KeepAspectRatio { get; set; }

        public override RtfCharFormat DefaultCharFormat => null;

        internal override string BlockHead
        {
            set { _blockHead = value; }
        }

        internal override string BlockTail
        {
            set { _blockTail = value; }
        }

        private string ExtractImage()
        {
            var result = new StringBuilder();

            for (var i = 0; i < _imgByte.Length; i++)
            {
                if (i != 0 && i % 60 == 0)
                    result.AppendLine();
                result.AppendFormat("{0:x2}", _imgByte[i]);
            }

            return result.ToString();
        }

        public override string Render()
        {
            var result = new StringBuilder(_blockHead);

            if (_startNewPage) result.Append(@"\pagebb");

            if (_margins[Direction.Top] >= 0) result.Append(@"\sb" + RtfUtility.Pt2Twip(_margins[Direction.Top]));
            if (_margins[Direction.Bottom] >= 0) result.Append(@"\sa" + RtfUtility.Pt2Twip(_margins[Direction.Bottom]));
            if (_margins[Direction.Left] >= 0) result.Append(@"\li" + RtfUtility.Pt2Twip(_margins[Direction.Left]));
            if (_margins[Direction.Right] >= 0) result.Append(@"\ri" + RtfUtility.Pt2Twip(_margins[Direction.Right]));
            switch (_alignment)
            {
                case Align.Left:
                    result.Append(@"\ql");
                    break;
                case Align.Right:
                    result.Append(@"\qr");
                    break;
                case Align.Center:
                    result.Append(@"\qc");
                    break;
            }
            result.AppendLine();

            result.Append(@"{\*\shppict{\pict");
            if (_imgType == ImageFileType.Jpg) result.Append(@"\jpegblip");
            else if (_imgType == ImageFileType.Png || _imgType == ImageFileType.Gif) result.Append(@"\pngblip");
            else throw new Exception("Image type not supported.");
            if (_height > 0) result.Append(@"\pichgoal" + RtfUtility.Pt2Twip(_height));
            if (_width > 0) result.Append(@"\picwgoal" + RtfUtility.Pt2Twip(_width));
            result.AppendLine();

            result.AppendLine(ExtractImage());
            result.AppendLine("}}");
            if (StartNewPara) result.Append(@"\par");
            result.AppendLine(_blockTail);
            return result.ToString();
        }
    }
}