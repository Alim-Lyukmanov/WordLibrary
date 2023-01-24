using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            WordTextPerformer wp = new WordTextPerformer(@"C:\Users\5\Desktop\WordTest\1.docx");
            wp.SetTextFont("Times New Roman");
            wp.CloseDoc();
            wp.CloseApp();
        }

    }
}
