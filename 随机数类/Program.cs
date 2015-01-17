using System;

namespace 随机数类
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("生产随机字符串："+BaseRandom.GetRandomString());
            Console.WriteLine("随机数：" + BaseRandom.GetRandom());
            Console.ReadKey();
        }
    }   
}
