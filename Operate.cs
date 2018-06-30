using Photoshop;
using System;
using System.IO;

namespace PsCore
{
    class Operate
    {
        public Application app = null;

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
        /// 保存PNG文件
        /// </summary>
        /// <param name="path"></param>
        public void SavePng(string path)
        {
            path = ExistPath(path);
            var docref = app.ActiveDocument;
            var name = docref.Name;
            var saveOptions = new PNGSaveOptions();
            saveOptions.Interlaced = false;
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
