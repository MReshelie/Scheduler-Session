using DevExpress.Utils.Drawing;
using DevExpress.Utils.Text;
using DevExpress.XtraSplashScreen;
using System;
using System.Drawing;

namespace Сессия.Classes
{
    class SplashImagePainter : ICustomImagePainter
    {
        ViewInfo info = null;

        static SplashImagePainter()
        {
            Painter = new SplashImagePainter();
        }

        protected SplashImagePainter() { }

        #region Drawing
        public void Draw(GraphicsCache cache, Rectangle bounds)
        {
            PointF pt = ViewInfo.CalcProgressLabelPoint(cache, bounds);
            cache.Graphics.DrawString(ViewInfo.Text, ViewInfo.ProgressLabelFont, ViewInfo.Brush, pt);
        }
        #endregion

        public static SplashImagePainter Painter { get; private set; }

        public ViewInfo ViewInfo
        {
            get
            {
                if (this.info == null) this.info = new ViewInfo();
                return this.info;
            }
        }
    }

    class ViewInfo
    {
        Brush brush = null;
        Font progressFont = null;

        public ViewInfo()
        {
            Counter = 0;
            Stage = string.Empty;
        }

        public PointF CalcProgressLabelPoint(GraphicsCache cache, Rectangle bounds)
        {
            Size size = TextUtils.GetStringSize(cache.Graphics, Text, ProgressLabelFont);
            return new Point(bounds.Width / 2 - size.Width / 2, bounds.Height - 40);
        }

        public Brush Brush
        {
            get
            {
                if (this.brush == null) this.brush = new SolidBrush(Color.DarkRed);
                return this.brush;
            }
        }

        public int Counter { get; set; }

        public Font ProgressLabelFont
        {
            get
            {
                if (this.progressFont == null) this.progressFont = new Font("Consolas", 12, FontStyle.Bold);
                return this.progressFont;
            }
        }

        public string Stage { get; set; }

        public string Text
        {
            get
            {
                if (Counter == 0)
                    return string.Format("{0}", Stage);
                else
                    return string.Format("{0} - ({1}%)", Stage, Counter.ToString("D2"));
            }
        }
    }
}