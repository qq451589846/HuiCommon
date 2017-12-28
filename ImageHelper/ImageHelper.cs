using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace Common
{
    public class ImageHelper
    {       

        /// <summary>
        /// 生成缩略图
        /// </summary>
        /// <param name="originalImagePath">源图路径（物理路径）</param>
        /// <param name="thumbnailPath">缩略图路径（物理路径）</param>
        /// <param name="thumbnailName">赋缩略图名字带后缀名</param>
        /// <param name="width">缩略图宽度</param>
        /// <param name="height">缩略图高度</param>
        public bool CreateThumbnail(string originalImagePath, string thumbnailPath, string thumbnailName, int width, int height)
        {
            if (!File.Exists(originalImagePath))
            {
                return false;
            }

            if (!Directory.Exists(thumbnailPath))
            {
                Directory.CreateDirectory(thumbnailPath);
            }

            thumbnailPath = Path.Combine(thumbnailPath, thumbnailName);

            if (File.Exists(thumbnailPath))
            {
                File.Delete(thumbnailPath);
            }
            Image originalImage = Image.FromFile(originalImagePath);

            try
            {
                if (originalImage.Width > width || originalImage.Height > height)
                {
                    Image newImage = CreateThumbnail(originalImage, height, width);
                    newImage.Save(thumbnailPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    newImage.Dispose();
                }
                else
                    originalImage.Save(thumbnailPath, ImageFormat.Jpeg);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                originalImage.Dispose();
            }
        }

        private Image CreateThumbnail(Image originalImage, int height, int width)
        {
            try
            {
                int towidth = width;
                int toheight = height;

                int x = 0;
                int y = 0;
                int ow = originalImage.Width;
                int oh = originalImage.Height;

                if (ow > oh)
                    toheight = originalImage.Height * width / originalImage.Width;
                else
                    towidth = originalImage.Width * height / originalImage.Height;

                //新建一个bmp图片
                Bitmap newImage = new Bitmap(towidth, toheight);

                //新建一个画板
                Graphics g = System.Drawing.Graphics.FromImage(newImage);

                //设置高质量插值法
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;

                //设置高质量,低速度呈现平滑程度
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

                //清空画布并以透明背景色填充
                g.Clear(Color.White);

                //在指定位置并且按指定大小绘制原图片的指定部分
                g.DrawImage(originalImage, new Rectangle(0, 0, towidth, toheight),
                 new Rectangle(x, y, ow, oh),
                 GraphicsUnit.Pixel);

                return newImage;
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {

            }
        }

        /// <summary>
        /// 二进制图片数据生成图片文件
        /// </summary>
        /// <param name="imgData"></param>
        /// <param name="PhyPath"></param>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public string CreateImageFile(byte[] imgData, string PhyPath, string FileName)
        {
            try
            {
                if (!Directory.Exists(PhyPath))
                    Directory.CreateDirectory(PhyPath);

                string FullName = Path.Combine(PhyPath, FileName);

                if (File.Exists(FullName))
                    File.Delete(FullName);

                Stream stream = new MemoryStream(imgData);

                Image img = Image.FromStream(stream);

                img.Save(FullName);

                return FullName;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private string CreateImageFile(byte[] imgData, string PhyPath, string FileName, int width, int height)
        {
            try
            {
                if (!Directory.Exists(PhyPath))
                    Directory.CreateDirectory(PhyPath);

                string FullName = Path.Combine(PhyPath, FileName);

                Stream stream = new MemoryStream(imgData);


                Image img = Image.FromStream(stream);

                Image newImage = CreateThumbnail(img, height, width);

                newImage.Save(FullName);

                return FullName;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }
    }
}
