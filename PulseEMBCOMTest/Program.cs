using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pulse.Extensions;
using Pulse;
using System.Diagnostics;
using System.Threading;
using System.IO;

namespace PulseEMBCOMTest
{
    class Program

    {
        public static IApplication PulseID;
        public static IEmbDesign myDesign;
        public static IBitmapImage myImage;
        public static IApplicationPool PulseIDPool;

        public static string CreateApplication()

        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            PulseIDPool = new ApplicationPool();
            PulseID = PulseIDPool.Allocate();
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            return elapsedTime;


        }


        public static string OpenTemplate (string templatePath)

        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            myDesign = PulseID.OpenDesign(templatePath, FileTypes.ftAuto, OpenTypes.otDefault, "Tajima");
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            return elapsedTime;
        }

        public static string SetText()

        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            if (myDesign.Elements.Count > 0)
            {
                IElement element1 = myDesign.Elements.Item[0];
                element1.ReplaceText("PULSE TEST");
              // System.Runtime.InteropServices.Marshal.ReleaseComObject(element1);

            }
           
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            return elapsedTime;
        }

        public static string RenderTemplate (string imagePath)

        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            Pulse.Extensions.RenderOptions ro = new Pulse.Extensions.RenderOptions();
            ro.Height = 300;
            ro.Width = 300;
            myImage = PulseEmbComExtensions.Render(myDesign, ro);
            myImage.Save(imagePath, ImageTypes.itAuto);
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            return elapsedTime;
        }

        public static string SaveTemplate(string outputPath)

        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            myDesign.Save(outputPath, FileTypes.ftAuto);
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            return elapsedTime;
        }


        static void Main(string[] args)
        {
            string[] fileList = Directory.GetFiles(args[0], "*.PXF");
            var csv = new StringBuilder();
            var headerLine = string.Format("{0},{1},{2},{3},{4},{5}","Design Name", "Create App", "Open Template", "Change Text", "Render Template", "Save Template");
            csv.AppendLine(headerLine);
            string createAppTime = CreateApplication();
            string version = PulseID.Version;
            foreach (string file in fileList)
            {

               
                string openTime = OpenTemplate(file);
                string changeElementTime = SetText();
                string imageFile = Path.Combine(args[1], Path.GetFileNameWithoutExtension(file)+".PNG");
                Console.WriteLine(imageFile);
                string outputFile = Path.Combine(args[1], Path.GetFileNameWithoutExtension(file)+".PXF");
                string renderTime = RenderTemplate(imageFile);
                string saveTemplateTime = SaveTemplate(outputFile);
                var newLine = string.Format("{0},{1},{2},{3},{4},{5}", Path.GetFileName(file), createAppTime, openTime, changeElementTime, renderTime, saveTemplateTime);
                csv.AppendLine(newLine);
                version = PulseID.Version;
              System.Runtime.InteropServices.Marshal.ReleaseComObject(myImage);
              System.Runtime.InteropServices.Marshal.ReleaseComObject(myDesign);
          
           

            }

            // System.Runtime.InteropServices.Marshal.ReleaseComObject(PulseID);
            Console.WriteLine(args[1]);
            File.WriteAllText(Path.Combine(args[1],"Timelog-"+version+".csv"), csv.ToString());

        }
    }
}
