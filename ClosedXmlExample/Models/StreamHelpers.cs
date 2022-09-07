using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXmlExample.Models
{
    public static class StreamHelpers
    {
        public static byte[] GetBytes(MemoryStream fs)
        {
            //MemoryStream fs = new MemoryStream();
            TextWriter tx = new StreamWriter(fs);

            tx.WriteLine("1111");
            tx.WriteLine("2222");
            tx.WriteLine("3333");

            tx.Flush();
            fs.Flush();

            fs.Position = 0;

            byte[] bytes = new byte[fs.Length];
            fs.Read(bytes, 0, bytes.Length);

            return bytes;
        }
    }
}
