using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Pipeline
{
    public abstract class Oportunidad
    {

        public const int HojaVariacion = 2;
        public const int HojaYTD = 3;
        public const int HojaYTG100 = 4;
        public const int HojaYTGPonderado = 5;
        public const int HojaPerdidas = 6;

        public int ColumnaCuenta;
        public int ColumnaOportunidad;
        public int ColumnaCodigo;
        public int ColumnaResponsable;
        public int ColumnaFase;
        public int ColumnaProbabilidad;
        public int ColumnaImporteUSD;
        public int ColumnaMonto;
        public int ColumnaPonderado;
        public int ColumnaFechaDeIngreso;

        public double Probabilidad { get; set; }
        public double Monto { get; set; }
        public double ImporteUSD { get; set; }

        public int Hoja { get; set; }
        public string Cuenta { get; set; }
        public string Nombre { get; set; }
        public int Codigo { get; set; }
        public string Responsable { get; set; }
        public string Fase { get; set; }
        public double Ponderado { get; set; }
        public string FechaDeIngreso { get; set; }


        public abstract void CargarDatos(ExcelWorksheet hoja, int i);

        public static Oportunidad CrearOportunidad(int hoja)
        {
            switch (hoja)
            {
                case Oportunidad.HojaYTD:
                    return new YTD();
                case Oportunidad.HojaYTG100:
                    return new YTG100();
                case Oportunidad.HojaYTGPonderado:
                    return new YTGPonderado();
                case Oportunidad.HojaPerdidas:
                    return new Perdidas();
            }
            return null;
        }
        public bool Iguales(Oportunidad otraOportunidad)
        {
            return otraOportunidad.Codigo == Codigo && otraOportunidad.Hoja == Hoja && otraOportunidad.Fase == Fase && ImporteUSD == otraOportunidad.ImporteUSD 
                && otraOportunidad.Probabilidad == Probabilidad && otraOportunidad.Monto == Monto && FechaDeIngreso == otraOportunidad.FechaDeIngreso;
        }
    }
}
