using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.UserSkins;
using DevExpress.Skins;
using DevExpress.LookAndFeel;
using System.Drawing;
using System.Reflection;
using DevExpress.XtraSplashScreen;
using Сессия.Classes;

namespace Сессия
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            BonusSkins.Register();
            SkinManager.EnableFormSkins();

            Image image = GetImage();
            // Show splash image
            SplashScreenManager.ShowImage(image, true, true, SplashImagePainter.Painter);
            SplashImagePainter.Painter.ViewInfo.Stage = "Загрузка приложения";
            Application.Run(new FormMain());
        }

        static Image GetImage()
        {
            Assembly asm = typeof(Program).Assembly;
            return new Bitmap(asm.GetManifestResourceStream("Сессия.Resources.rasp.jpg"));
        }
    }
}
