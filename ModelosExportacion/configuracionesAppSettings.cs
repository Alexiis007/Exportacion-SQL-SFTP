using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelosExportacion
{
    public  class configuracionesAppSettings
    {
        public  string bdcnServer { set; get; }
        public  string bdcnBD { set; get; }
        public  string bdcnUsuario { set; get; }
        public  string bdcnContraseña { set; get; }

        //public string bdspSLValida { set; get; }

        public  string sftpcnServer_ICMMNFHeinekenQA_SFTP { set; get; }
        public  string sftpcnPuerto_ICMMNFHeinekenQA_SFTP { set; get; }
        public  string sftpcnUsuario_ICMMNFHeinekenQA_SFTP { set; get; }
        public  string sftpcnContraseña_ICMMNFHeinekenQA_SFTP { set; get; }

        public  string sftpcnServer_ICMCOMHeinekenQA_SFTP { set; get; }
        public  string sftpcnPuerto_ICMCOMHeinekenQA_SFTP { set; get; }
        public  string sftpcnUsuario_ICMCOMHeinekenQA_SFTP { set; get; }
        public  string sftpcnContraseña_ICMCOMHeinekenQA_SFTP { set; get; }

        public  string sftpcnServer_ICMCOMCatorcenalHeinekenQA_SFTP { set; get; }
        public  string sftpcnPuerto_ICMCOMCatorcenalHeinekenQA_SFTP { set; get; }
        public  string sftpcnUsuario_ICMCOMCatorcenalHeinekenQA_SFTP { set; get; }
        public  string sftpcnContraseña_ICMCOMCatorcenalHeinekenQA_SFTP { set; get; }


        public  string strTablasExportar { set; get; }

        public  int intMaximoRegistros { set; get; }
        public  string strIronLicenseKey { set; get; }

        public  string dteFechaFiltro { set; get; }

        public  int intMiliSegundosEspera { set; get; }

        public  string ubiRutaCarpetaLocalExcel { set; get; }
        public  string ubiRutaCarpetaDestinoExcel { set; get; }

        public  string ubiRutaCarpetaDestinoTablasToSQL { set; get; }
        public  string  ubiRutaCarpetaOrigenTablasToSQL { set; get; }

    }
}
