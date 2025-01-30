using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelosExportacion
{
    public class Encoder
    {
        public string EncriptarBase64(string _cadenaAencriptar)
        {
            return Convert.ToBase64String(Encoding.Unicode.GetBytes(_cadenaAencriptar));
        }

        public string DesEncriptarBase64(string _cadenaAdesencriptar)
        {
            return Encoding.Unicode.GetString(Convert.FromBase64String(_cadenaAdesencriptar));
        }


    }
}
