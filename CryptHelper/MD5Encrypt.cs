using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.IO;
using System.Web.Security;

namespace Common.CryptHelper
{
    /// <summary>
    /// MD5加密类,注意经MD5加密过的信息是不能转换回原始数据的
    /// ,请不要在用户敏感的信息中使用此加密技术,比如用户的密码,
    /// 请尽量使用对称加密
    /// </summary>
    public class MD5Encrypt
    {
        private MD5 md5;
        /// <summary>
        /// 构造函数
        /// </summary>
        public MD5Encrypt()
        {
            md5 = new MD5CryptoServiceProvider();
        }

        /// <summary>
        /// 32位MD5加密
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string Md5Hash(string input)
        {
            MD5CryptoServiceProvider md5Hasher = new MD5CryptoServiceProvider();
            byte[] data = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input));
            StringBuilder sBuilder = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            return sBuilder.ToString();
        }

        /// <summary>
        /// 从字符串中获取散列值
        /// </summary>
        /// <param name="str">要计算散列值的字符串</param>
        /// <returns></returns>
        public string GetMD5FromString(string str)
        {
            byte[] toCompute = Encoding.Unicode.GetBytes(str);
            byte[] hashed = md5.ComputeHash(toCompute, 0, toCompute.Length);
            return Encoding.ASCII.GetString(hashed);
        }
        /// <summary>
        /// 根据文件来计算散列值
        /// </summary>
        /// <param name="filePath">要计算散列值的文件路径</param>
        /// <returns></returns>
        public string GetMD5FromFile(string filePath)
        {
            bool isExist = File.Exists(filePath);
            if (isExist)//如果文件存在
            {
                FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                StreamReader reader = new StreamReader(stream, Encoding.Unicode);
                string str = reader.ReadToEnd();
                byte[] toHash = Encoding.Unicode.GetBytes(str);
                byte[] hashed = md5.ComputeHash(toHash, 0, toHash.Length);
                stream.Close();
                return Encoding.ASCII.GetString(hashed);
            }
            else//文件不存在
            {
                throw new FileNotFoundException("File not found!");
            }
        }
    }

    /// <summary>
    /// MD5
    /// 单向加密
    /// </summary>
    public class EncryptionMD5
    {
        /// <summary>
        /// 获得一个字符串的加密密文
        /// 此密文为单向加密，即不可逆(解密)密文
        /// </summary>
        /// <param name="plainText">待加密明文</param>
        /// <returns>已加密密文</returns>
        public static string EncryptString(string plainText)
        {
            return EncryptStringMD5(plainText);
        }
        /// <summary>
        /// 获得一个字符串的加密密文
        /// 此密文为单向加密，即不可逆(解密)密文
        /// </summary>
        /// <param name="plainText">待加密明文</param>
        /// <returns>已加密密文</returns>
        public static string EncryptStringMD5(string plainText)
        {
            string encryptText = "";
            if (string.IsNullOrEmpty(plainText)) return encryptText;
            encryptText = FormsAuthentication.HashPasswordForStoringInConfigFile(plainText, "md5");
            return encryptText;
        }
        /// <summary>
        /// 判断明文与密文是否相符
        /// </summary>
        /// <param name="plainText">待检查的明文</param>
        /// <param name="encryptText">待检查的密文</param>
        /// <returns>bool</returns>
        public static bool EqualEncryptString(string plainText, string encryptText)
        {
            return EqualEncryptStringMD5(plainText, encryptText);
        }
        /// <summary>
        /// 判断明文与密文是否相符
        /// </summary>
        /// <param name="plainText">待检查的明文</param>
        /// <param name="encryptText">待检查的密文</param>
        /// <returns>bool</returns>
        public static bool EqualEncryptStringMD5(string plainText, string encryptText)
        {
            bool result = false;
            if (string.IsNullOrEmpty(plainText) || string.IsNullOrEmpty(encryptText))
                return result;
            result = EncryptStringMD5(plainText).Equals(encryptText);
            return result;
        }
    }
    /// <summary>
    /// SHA1
    /// 单向加密
    /// </summary>
    public class EncryptionSHA1
    {
        /// <summary>
        /// 获得一个字符串的加密密文
        /// 此密文为单向加密，即不可逆(解密)密文
        /// </summary>
        /// <param name="plainText">待加密明文</param>
        /// <returns>已加密密文</returns>
        public static string EncryptString(string plainText)
        {
            return EncryptStringSHA1(plainText);
        }
        /// <summary>
        /// 获得一个字符串的加密密文
        /// 此密文为单向加密，即不可逆(解密)密文
        /// </summary>
        /// <param name="plainText">待加密明文</param>
        /// <returns>已加密密文</returns>
        public static string EncryptStringSHA1(string plainText)
        {
            string encryptText = "";
            if (string.IsNullOrEmpty(plainText)) return encryptText;
            encryptText = FormsAuthentication.HashPasswordForStoringInConfigFile(plainText, "sha1");
            return encryptText;
        }
        /// <summary>
        /// 判断明文与密文是否相符
        /// </summary>
        /// <param name="plainText">待检查的明文</param>
        /// <param name="encryptText">待检查的密文</param>
        /// <returns>bool</returns>
        public static bool EqualEncryptString(string plainText, string encryptText)
        {
            return EqualEncryptStringSHA1(plainText, encryptText);
        }
        /// <summary>
        /// 判断明文与密文是否相符
        /// </summary>
        /// <param name="plainText">待检查的明文</param>
        /// <param name="encryptText">待检查的密文</param>
        /// <returns>bool</returns>
        public static bool EqualEncryptStringSHA1(string plainText, string encryptText)
        {
            bool result = false;
            if (string.IsNullOrEmpty(plainText) || string.IsNullOrEmpty(encryptText))
                return result;
            result = EncryptStringSHA1(plainText).Equals(encryptText);
            return result;
        }
    }
}
