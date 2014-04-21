using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Pipeline
{

    public class YTD :Oportunidad
    {
        public YTD()
        {
            ColumnaCuenta = 2;
            ColumnaOportunidad = 3;
            ColumnaCodigo = 4;
            ColumnaResponsable = 5;
            ColumnaFase = 6;
            ColumnaImporteUSD = 7;
            ColumnaPonderado = 8;
            Hoja = HojaYTD;
        }

        public override double Probabilidad()
        {
            return 1;
        }

        public override double Monto()
        {
            return Ponderado;
        }


        public override void CargarDatos(ExcelWorksheet hoja, int i)
        {
            Cuenta = hoja.GetValue<string>(i, ColumnaCuenta);
            Nombre = hoja.GetValue<string>(i, ColumnaOportunidad);
            Codigo = hoja.GetValue<int>(i, ColumnaCodigo);
            Responsable = hoja.GetValue<string>(i, ColumnaResponsable);
            Fase = hoja.GetValue<string>(i, ColumnaFase);
            ImporteUSD(hoja.GetValue<double>(i, ColumnaImporteUSD));
            Ponderado = hoja.GetValue<double>(i, ColumnaPonderado);
        }
    }
}
