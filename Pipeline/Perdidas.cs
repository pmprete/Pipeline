using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Pipeline
{
    public class Perdidas: Oportunidad
    {

        public Perdidas()
        {
            ColumnaCuenta = 2;
            ColumnaOportunidad = 3;
            ColumnaCodigo = 4;
            ColumnaResponsable = 5;
            ColumnaFase = 6;
            ColumnaPonderado = 7;

            Hoja = HojaPerdidas;
            Ponderado = 0;
        }

        public override double Probabilidad()
        {
            return 0;
        }

        public override double Monto()
        {
            return 0;
        }

        public override double ImporteUSD()
        {
            return 0;
        }

        public override void CargarDatos(ExcelWorksheet hoja, int i)
        {
            Cuenta = hoja.GetValue<string>(i, ColumnaCuenta);
            Nombre = hoja.GetValue<string>(i, ColumnaOportunidad);
            Codigo = hoja.GetValue<int>(i, ColumnaCodigo);
            Responsable = hoja.GetValue<string>(i, ColumnaResponsable);
            Fase = hoja.GetValue<string>(i, ColumnaFase);
            Ponderado = hoja.GetValue<double>(i, ColumnaPonderado);
        }

      
    }
}
