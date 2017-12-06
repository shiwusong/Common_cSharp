using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImgtoString
{
    /// <summary>
    /// 对图片的一些操作
    /// </summary>
    class ImgClass
    {
        /// <summary>
        /// 获取bitmap对象
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public Bitmap GetImg(string filename)
        {
            FileStream fs = File.OpenRead(filename); //OpenRead
            int filelength = 0;
            filelength = (int)fs.Length; //获得文件长度 
            Byte[] image = new Byte[filelength]; //建立一个字节数组 
            fs.Read(image, 0, filelength); //按字节流读取 
            System.Drawing.Image result = System.Drawing.Image.FromStream(fs);
            fs.Close();
            Bitmap bit = new Bitmap(result);
            return bit;
        }

        /// <summary>
        /// 将bitmap对象转化为灰度二维数组
        /// </summary>
        /// <param name="bp"></param>
        /// <param name="isAverage"></param>
        /// <returns></returns>
        public int[,] ImgtoArray(Bitmap bp, bool isAverage = false)
        {
            int[,] imgArr = new int[bp.Height, bp.Width];
            for (int i = 0; i < bp.Height; i++)
            {
                for (int j = 0; j < bp.Width; j++)
                {
                    System.Drawing.Color c = bp.GetPixel(i, j);
                    if (isAverage)
                    {
                        imgArr[i, j] = (c.R + c.G + c.B) / 3;
                    }
                    else
                    {
                        //imgArr[i, j] = (int)( (255 -c.R) * 0.3 + (255 - c.G) * 0.59 + (255 - c.B) * 0.11);
                        imgArr[i, j] = (int)((c.R) * 0.3 + (c.G) * 0.59 + (c.B) * 0.11);
                    }
                }
            }
            return imgArr;
        }

        /// <summary>
        /// 返回减小规模的二维数组（取平均值）
        /// </summary>
        /// <param name="arr">原数组</param>
        /// <param name="height">缩减后的高度</param>
        /// <param name="width">缩减后的宽度</param>
        /// <returns></returns>
        public int[,] DevArray(int[,] arr, int height, int width)
        {
            int arrH = arr.GetLength(0);//原数组高度
            int arrW = arr.GetLength(1);//原数组宽度
            height = height >= arrH ? arrH : height; //缩减后高度
            width = width >= arrW ? arrW : width;     //缩减后宽度
            int[,] devArr = new int[height, width];
            int nH = arrH / height;  //高度上分的间隔数
            int nW = arrW / width;   //宽度上分的间隔数
            for (int i0 = 0; i0 < height; i0++)
            {
                for (int j0 = 0; j0 < width; j0++)
                {
                    int average = 0;
                    for (int i1 = i0 * nH; i1 < i0 * nH + nH; i1++)
                    {
                        for (int j1 = j0 * nW; j1 < j0 * nW + nW; j1++)
                        {
                            average += arr[i1, j1];
                        }
                    }
                    average /= nH * nW;
                    devArr[i0, j0] = average;
                }
            }
            return devArr;
        }

        public void print(int[,] arr)
        {
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    Console.Write(arr[i, j] + " ");
                }
                Console.WriteLine();
            }
        }

        public void print(char[,] arr)
        {
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    Console.Write(arr[i, j] + " ");
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// 把arr数组转化为char数组
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        public char[,] ArrtoChar(int[,] arr)
        {
            string lib = @"$@B%8&WM#*oahkbdpqwmZO0QLCJUYXzcvunxrjft/\|()1{}[]?-_+~<>i!lI;:,\^`'.       ";
            int len = lib.Length;

            int height = arr.GetLength(0);
            int width = arr.GetLength(1);
            char[,] charArr = new char[width, height];
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    int index = (int)(arr[i, j] / 255.0 * len); //建立映射
                    index = index >= len ? len - 1 : index;
                    charArr[j, i] = lib[index];
                }
            }
            return charArr;
        }
    }
}
