using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common.ZipHelper
{
    public class ZipHelper
    {
        /// <summary>
        /// 压缩字节
        /// 1.创建压缩的数据流 
        /// 2.设定compressStream为存放被压缩的文件流,并设定为压缩模式
        /// 3.将需要压缩的字节写到被压缩的文件流 
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static byte[] CompressBytes(byte[] bytes)
        {
            using (MemoryStream compressStream = new MemoryStream())
            {
                using (var zipStream = new GZipStream(compressStream, CompressionMode.Compress))
                    zipStream.Write(bytes, 0, bytes.Length);
                return compressStream.ToArray();
            }
        }

        /// <summary>
        /// 解压缩字节
        /// 1.创建被压缩的数据流
        /// 2.创建zipStream对象，并传入解压的文件流
        /// 3.创建目标流
        /// 4.zipStream拷贝到目标流
        /// 5.返回目标流输出字节
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static byte[] Decompress(byte[] bytes)
        {
            using (var compressStream = new MemoryStream(bytes))
            {
                using (var zipStream = new GZipStream(compressStream, CompressionMode.Decompress))
                {
                    using (var resultStream = new MemoryStream())
                    {
                        zipStream.CopyTo(resultStream);
                        return resultStream.ToArray();
                    }
                }
            }
        }

        public static void ZipFile(string fileToZip, string zipedFile, int compressionLevel, int blockSize)
        {
            if (!File.Exists(fileToZip))
            {
                throw new FileNotFoundException("指定要压缩的文件: " + fileToZip + " 不存在!");
            }
            using (FileStream fileStream = File.Create(zipedFile))
            {
                using (ZipOutputStream zipOutputStream = new ZipOutputStream(fileStream))
                {
                    using (FileStream fileStream2 = new FileStream(fileToZip, FileMode.Open, FileAccess.Read))
                    {
                        string name = fileToZip.Substring(fileToZip.LastIndexOf("\\") + 1);
                        ZipEntry entry = new ZipEntry(name);
                        zipOutputStream.PutNextEntry(entry);
                        zipOutputStream.SetLevel(compressionLevel);
                        byte[] array = new byte[blockSize];
                        int num = 0;
                        try
                        {
                            do
                            {
                                num = fileStream2.Read(array, 0, array.Length);
                                zipOutputStream.Write(array, 0, num);
                            }
                            while (num > 0);
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        fileStream2.Close();
                    }
                    zipOutputStream.Finish();
                    zipOutputStream.Close();
                }
                fileStream.Close();
            }
        }

        public static void ZipFile(string fileToZip, string zipedFile)
        {
            if (!File.Exists(fileToZip))
            {
                throw new FileNotFoundException("指定要压缩的文件: " + fileToZip + " 不存在!");
            }
            using (FileStream fileStream = File.OpenRead(fileToZip))
            {
                byte[] array = new byte[fileStream.Length];
                fileStream.Read(array, 0, array.Length);
                fileStream.Close();
                using (FileStream baseOutputStream = File.Create(zipedFile))
                {
                    using (ZipOutputStream zipOutputStream = new ZipOutputStream(baseOutputStream))
                    {
                        string name = fileToZip.Substring(fileToZip.LastIndexOf("\\") + 1);
                        ZipEntry entry = new ZipEntry(name);
                        zipOutputStream.PutNextEntry(entry);
                        zipOutputStream.SetLevel(5);
                        zipOutputStream.Write(array, 0, array.Length);
                        zipOutputStream.Finish();
                        zipOutputStream.Close();
                    }
                }
            }
        }

        public static void ZipFileDirectory(DirectoryInfo dir, string zipedFile, string password)
        {
            using (FileStream baseOutputStream = File.Create(zipedFile))
            {
                using (ZipOutputStream zipOutputStream = new ZipOutputStream(baseOutputStream))
                {
                    if (password != null && password.Length > 0)
                    {
                        zipOutputStream.Password = password;
                    }
                    zipOutputStream.SetLevel(9);
                    ZipHelper.ZipSetp(dir.FullName, zipOutputStream, "");
                }
            }
        }

        private static void ZipSetp(string strDirectory, ZipOutputStream s, string parentPath)
        {
            if (strDirectory[strDirectory.Length - 1] != Path.DirectorySeparatorChar)
            {
                strDirectory += Path.DirectorySeparatorChar.ToString();
            }
            Crc32 crc = new Crc32();
            string[] fileSystemEntries = Directory.GetFileSystemEntries(strDirectory);
            string[] array = fileSystemEntries;
            foreach (string text in array)
            {
                if (Directory.Exists(text))
                {
                    string str = parentPath + text.Substring(text.LastIndexOf("\\") + 1);
                    str += "\\";
                    ZipHelper.ZipSetp(text, s, str);
                }
                else
                {
                    using (FileStream fileStream = File.OpenRead(text))
                    {
                        byte[] array2 = new byte[fileStream.Length];
                        fileStream.Read(array2, 0, array2.Length);
                        string name = parentPath + text.Substring(text.LastIndexOf("\\") + 1);
                        ZipEntry zipEntry = new ZipEntry(name);
                        zipEntry.DateTime = DateTime.Now;
                        zipEntry.Size = fileStream.Length;
                        fileStream.Close();
                        crc.Reset();
                        crc.Update(array2);
                        zipEntry.Crc = crc.Value;
                        s.PutNextEntry(zipEntry);
                        s.Write(array2, 0, array2.Length);
                    }
                }
            }
        }

        public static void UnZip(string zipedFile, string strDirectory, string password, bool overWrite)
        {
            if (strDirectory == "")
            {
                strDirectory = Directory.GetCurrentDirectory();
            }
            if (!strDirectory.EndsWith("\\"))
            {
                strDirectory += "\\";
            }
            using (ZipInputStream zipInputStream = new ZipInputStream(File.OpenRead(zipedFile)))
            {
                zipInputStream.Password = password;
                ZipEntry nextEntry;
                while ((nextEntry = zipInputStream.GetNextEntry()) != null)
                {
                    string str = "";
                    string text = "";
                    text = nextEntry.Name;
                    if (text != "")
                    {
                        str = Path.GetDirectoryName(text) + "\\";
                    }
                    string fileName = Path.GetFileName(text);
                    Directory.CreateDirectory(strDirectory + str);
                    if (fileName != "" && ((File.Exists(strDirectory + str + fileName) & overWrite) || !File.Exists(strDirectory + str + fileName)))
                    {
                        using (FileStream fileStream = File.Create(strDirectory + str + fileName))
                        {
                            int num = 2048;
                            byte[] array = new byte[2048];
                            while (true)
                            {
                                num = zipInputStream.Read(array, 0, array.Length);
                                if (num <= 0)
                                {
                                    break;
                                }
                                fileStream.Write(array, 0, num);
                            }
                            fileStream.Close();
                        }
                    }
                }
                zipInputStream.Close();
            }
        }
    }
}
