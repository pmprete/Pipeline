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

        protected double _probabilidad;
        protected double _monto;
        protected double _importeUSD;

        public int Hoja { get; set; }
        public string Cuenta { get; set; }
        public string Nombre { get; set; }
        public int Codigo { get; set; }
        public string Responsable { get; set; }
        public string Fase { get; set; }
        public double Ponderado { get; set; }

        public virtual double ImporteUSD()
        {
            return _importeUSD;
        }

        public void ImporteUSD(double value)
        {
            _importeUSD = value;
        }

        public virtual double Probabilidad()
        {
            return _probabilidad;
        }

        public void Probabilidad(double value)
        {
             _probabilidad = value; 
        }


        public virtual double Monto()
        {
            return _monto;
        }

        public void Monto(double value)
        {
            _monto = value;
        }

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

    }
}
