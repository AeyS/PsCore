using System;
using System.IO;
using Photoshop;

namespace PsCore
{
    class Program : Operate
    {
        public const int PSD = 1;
        public const int JPG = 2;
        public const int PNG = 3;
        public const double px750 = 750;
        public const double px800 = 800;
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
                        m = psoft.InputNum("1.PSD 2.JPG 3.PNG");
                        psoft.EachSaveFile(m);
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

        public void EachSaveFile(int format)
        {
            foreach (Document doc in app.Documents)
            {
                app.ActiveDocument = doc;
                SaveFile(format);
            }
        }

        public void SaveFile(int format)
        {
            string path = Path.GetFullPath(app.ActiveDocument.Path);
            Console.WriteLine(path);
            switch (format)
            {
                case JPG:
                    SaveJpg(path);
                    break;
                case PNG:
                    SavePng(path);
                    break;
                case PSD:
                    SavePsd(path);
                    break;
                default:
                    Console.WriteLine("没有指定保存格式！");
                    break;
            }
        }
    }
}
