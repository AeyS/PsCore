using System;
using System.IO;
using Photoshop;

namespace PsCore
{
    class Program
    {
        public Application app = null;
        public string JPG = "jpg";
        public string PSD = "psd";
        public double px750 = 750;
        public double px800 = 800;
        public string abspath = "";

        /// <summary>
        /// 功能菜单选择
        /// </summary>
        /// <param name="showInfo"></param>
        /// <returns></returns>
        public int InputNum(string showInfo)
        {
            Console.WriteLine(showInfo);
            int m = 0;
            try
            {
                m = int.Parse(Console.ReadLine());
            }
            catch
            {
                Console.WriteLine("输入错误，请输入功能序号！");
                return InputNum(showInfo);
            }
            return m;
        }

        static void Main(string[] args)
        {
            Program psoft = new Program();
            psoft.SetUp();
            Console.WriteLine("-----------版本号：201707061347----------");
            while (true)
            {
                int m = psoft.InputNum("1.京东主图\n2.全部保存");
                switch(m)
                {
                    case 1:
                        psoft.JdMainPic();
                        break;
                    case 2:
                        m = psoft.InputNum("1.PSD 2.JPG");
                        if (m == 1)
                        {
                            psoft.EachSaveFile(psoft.PSD);
                        }
                        else
                        {
                            psoft.EachSaveFile(psoft.JPG);
                        }
                        
                        break;
                }
            }
        }

        public void SetUp()
        {
            Console.WriteLine("尝试连接Photoshop软件...");
            if (app == null)
            {
                Console.WriteLine("正在启动软件...");
                app = new Application();
            }
            Console.WriteLine("已连接！");
            Console.Clear();
        }

        public void SetPixels()
        {
            // 将单位设置成px
            app.Preferences.RulerUnits = PsUnits.psPixels;
        }

        /// <summary>
        /// 京东主图
        /// </summary>
        public void JdMainPic()
        {
            this.SetPixels();
            foreach (Document doc in app.Documents)
            {
                app.ActiveDocument = doc;
                ResizeAndCanvas(px800);
            }
        }

        public void EachSaveFile(string format)
        {
            foreach (Document doc in app.Documents)
            {
                app.ActiveDocument = doc;
                SaveFile(format);
            }
        }

        public void SaveFile(string format)
        {
            string path = Path.GetFullPath(app.ActiveDocument.Path);
            Console.WriteLine(path);
            if(format == JPG)
            {
                SaveJpg(path);
            }
            else if (format == PSD)
            {
                SavePsd(path);
            }
            else
            {
                Console.WriteLine("没有指定保存格式！");
            }
        }

        public string ExistPath(string path)
        {
            // 的那个路径结尾没有/符号，则补上
            if (path.Substring(path.Length - 1) != "/" && path.Substring(path.Length - 1) != @"\")
                path = path + "/";
            // 如果不存在，则创建路径
            if (Directory.Exists(path) != true)
                Directory.CreateDirectory(path);
            return path;
        }

        /// <summary>
        /// 保存PSD文件
        /// </summary>
        /// <param name="path"></param>
        public void SavePsd(string path)
        {
            path = ExistPath(path);
            var docref = app.ActiveDocument;
            var name = docref.Name;
            var saveOptions = new PhotoshopSaveOptions();
            saveOptions.AlphaChannels = false;
            saveOptions.Annotations = false;
            saveOptions.Layers = true;
            saveOptions.SpotColors = false;
            docref.SaveAs(path + name, saveOptions, false, PsExtensionType.psLowercase);
        }

        /// <summary>
        /// 保存JPG文件
        /// </summary>
        /// <param name="path"></param>
        public void SaveJpg(string path)
        {
            path = ExistPath(path);
            var docref = app.ActiveDocument;
            var name = docref.Name;
            var saveOptions = new JPEGSaveOptions();
            saveOptions.EmbedColorProfile = true;
            saveOptions.Quality = 10;
            Console.WriteLine(path + name);
            docref.SaveAs(path + name, saveOptions, false, PsExtensionType.psLowercase);
        }

        /// <summary>
        /// 批量重置图片大小
        /// </summary>
        /// <param name="pxWidth"></param>
        public void ResizeAndCanvas(double pxWidth)
        {
            // 获取图片宽高
            double width = app.ActiveDocument.Width;
            double height = app.ActiveDocument.Height;
            // 存放图片最大值
            double deuce = 0;
            // 当图片最大值找到后，将其赋值给deuce，然后将图片的宽高值都改成deuce
            deuce = width > height ? width : height;
            Console.WriteLine("width:{0}, height:{1}", width, height);
            ResizeCanvas(deuce);
            ResizeImage(pxWidth);
        }

        /// <summary>
        /// 重置画布大小
        /// </summary>
        /// <param name="deuce"></param>
        public void ResizeCanvas(double deuce)
        {
            ResizeCanvas(deuce, deuce);
        }

        /// <summary>
        /// 重置画布大小
        /// </summary>
        /// <param name="width"></param>
        /// <param name="heigth"></param>
        public void ResizeCanvas(double width, double heigth)
        {
            app.ActiveDocument.ResizeCanvas(width, heigth, null);
        }

        /// <summary>
        /// 重置图片分辨率
        /// </summary>
        /// <param name="deuce"></param>
        public void ResizeImage(double deuce)
        {
            this.ResizeImage(deuce, deuce);
        }

        /// <summary>
        /// 重置图片分辨率
        /// </summary>
        /// <param name="widht"></param>
        /// <param name="height"></param>
        public void ResizeImage(double widht, double height)
        {
            app.ActiveDocument.ResizeImage(widht, height, null, PsResampleMethod.psBicubic);
        }
    }
}
