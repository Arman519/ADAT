using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Lync.Model;

namespace DesktopDevDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var theClient = LyncClient.GetClient();
            Console.WriteLine(theClient.State);

            Console.ReadLine();
        }
    }
}
