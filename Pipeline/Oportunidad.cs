using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Pipeline
{
    public abstract class Oportunidad
    {

        public const string HojaVariacion = "Variación";
        public const string HojaYTD = "YTD";
        public const string HojaYTG100 = "YTG 100%";
        public const string HojaYTGPonderado = "YTG Ponderado";
        public const string HojaPerdidas = "Opps perdidas";

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

        public string Hoja { get; set; }
        public string Cuenta { get; set; }
        public string Nombre { get; set; }
        public int Codigo { get; set; }
        public string Responsable { get; set; }
        public string Fase { get; set; }
        public double Ponderado { get; set; }
        public DateTime FechaDeIngreso { get; set; }


        public abstract void CargarDatos(ExcelWorksheet hoja, int i);

        public static Oportunidad CrearOportunidad(string hoja)
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
            return otraOportunidad.Codigo == Codigo && otraOportunidad.Hoja == Hoja && otraOportunidad.Fase.Trim() == Fase.Trim() && ImporteUSD == otraOportunidad.ImporteUSD
                && otraOportunidad.Nombre.Trim() == Nombre.Trim() && otraOportunidad.Probabilidad == Probabilidad && otraOportunidad.Monto == Monto 
                && DateTime.Compare(FechaDeIngreso, otraOportunidad.FechaDeIngreso) == 0;
        }

        public DateTime ConvertirExcelAFecha(ExcelWorksheet hoja, int i, int columnaFechaDeIngreso)
        {
            hoja.Cells[i, columnaFechaDeIngreso].Style.Numberformat.Format = "#,##0";
            var textoFecha = hoja.Cells[i, columnaFechaDeIngreso].Text;
            textoFecha = String.IsNullOrEmpty(textoFecha) || textoFecha == "Ninguno" ? "0" : textoFecha;
            var fechaNumero = Convert.ToDouble(textoFecha);
            var fecha = DateTime.FromOADate(fechaNumero);
            hoja.Cells[i, columnaFechaDeIngreso].Style.Numberformat.Format = "DD/MM/YYYY";
            return fecha;
        }
    }
}
