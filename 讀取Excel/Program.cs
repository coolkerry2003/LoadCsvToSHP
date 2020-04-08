using PilotGaea.Geometry;
using PilotGaea.TMPEngine;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 讀取Excel
{
    class Program
    {
        //定義MapServer==================================================
        private const string mTmpxPath = @"C:\ProgramData\PilotGaea\PGMaps\LandGen\Map.TMPX";
        private const string mPluginDir = "";
        private const int mMapPort = 8888;
        //===============================================================

        //定義OleDb======================================================
        //1.檔案位置    注意絕對路徑 -> 非 \  是 \\
        private const string FileName = @"C:\Users\User\Desktop\控制點\104-108基隆市加密控制點.csv";
        //2.第一行是否為標題
        private const bool Hdr = false;
        //3.欄位
        private static List<string> Field = new List<string>();
        //4.內容
        private static List<string[]> Content = new List<string[]>();
        //5.座標索引
        private static int posX = 1, posY = 2;
        //6.預設產出位置
        private static string Path = AppDomain.CurrentDomain.BaseDirectory + "Gen.SHP";
        //===============================================================
        static void Main(string[] args)
        {
            Console.WriteLine("開始進行Excel轉檔");
            
            Console.WriteLine("讀取csv...");
            FileLoad();

            Console.WriteLine("開始建立圖檔...");
            Gen();
        }
        /// <summary>
        /// 讀取csv
        /// </summary>
        static void FileLoad()
        {
            StreamReader sr = new StreamReader(FileName, Encoding.Default);
            string line = string.Empty;
            int i = 0;
            while((line = sr.ReadLine()) != null)
            {
                if (Hdr && i == 0)
                {
                    Field = line.Split(',').ToList();
                }
                else if (!Hdr && i == 0)
                {
                    int length = line.Split(',').Length;
                    Field = new string[length].ToList();
                    for (int j = 0; j < length; j++)
                    {
                        Field[j] = (Convert.ToChar(65 + j).ToString());
                    }
                }
                else
                {
                    string[] cont = line.Split(',');
                    if (Field.Count == cont.Length)
                    {
                        Content.Add(cont);
                    }
                    else
                    {
                        Console.WriteLine("不符長度" + string.Concat(cont));
                    }
                }
                i++;
            }
        }
        static void Gen()
        {
            CServer m_Server = new CServer();
            if (m_Server.Init(mMapPort, mTmpxPath, mPluginDir))
            {
                Console.WriteLine("初始化成功!");

                Console.WriteLine("建立Shp...");
                CGeoDatabase db = m_Server.GeoDB;
                CSHPFile m_SHP = db.CreateSHPFile();

                //初始化欄位、型態、長度
                List<string> FieldNames = new List<string>();
                List<FIELD_TYPE> FieldTypes = new List<FIELD_TYPE>();
                List<int> FieldLengths = new List<int>();

                //建立欄位
                Console.WriteLine("建立欄位...");
                foreach (string s in Field)
                {
                    FieldNames.Add(s);
                    FieldTypes.Add(FIELD_TYPE.STRING);
                    FieldLengths.Add(255);
                }
                m_SHP.Create(FieldNames.ToArray(), FieldTypes.ToArray(), FieldLengths.ToArray());

                //建置圖素
                Console.WriteLine("建置圖素...");
                foreach (string[] strs in Content)
                {
                    GeoPoint p = new GeoPoint();
                    p.x = Convert.ToDouble(strs[posX]);
                    p.y = Convert.ToDouble(strs[posY]);
                    m_SHP.CreateEntity(p, strs);
                }
                m_SHP.Save(Path);
                m_SHP.Close();
                Console.WriteLine("產製完畢!");

            }
            else
            {
                Console.WriteLine("初始化Sevrer失敗...");
            }
            
        }
    }
}
