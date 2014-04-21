using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Pipeline
{
    public class Variacion: Oportunidad
    {
        public int ColumnaFaseAnterior;
        public int ColumnaProbabilidadAnterior;
        public int ColumnaImporteUSDAnterior;
        public int ColumnaMontoAnterior;
        public int ColumnaPonderadoAnterior;

        public int ColumnaValidacionYTD;
        public int ColumnaValidacionYtg100;
        public int ColumnaValidacionYtgPonderado;
        public int ColumnaValidacionTc;
        public int ColumnaValidacionDiferencia;

        public string FaseAnterior { get; set; }
        public double ProbabilidadAnterior { get; set; }
        public double ImporteUSDAnterior { get; set; }
        public double MontoAnterior { get; set; }
        public double PonderadoAnterior { get; set; }
        public int HojaActual { get; set; }
        public int HojaAnterior { get; set; }


        public Variacion()
        {
            ColumnaCuenta = 2;
            ColumnaOportunidad = 3;
            ColumnaCodigo = 4;
            ColumnaResponsable = 5;
            ColumnaFase = 6;
            ColumnaFaseAnterior = 7;
            ColumnaProbabilidad = 8;
            ColumnaProbabilidadAnterior = 9;
            ColumnaImporteUSD = 10;
            ColumnaImporteUSDAnterior = 11;
            ColumnaMonto = 12;
            ColumnaMontoAnterior = 13;
            ColumnaPonderado = 14;
            ColumnaPonderadoAnterior = 15;
            ColumnaValidacionYTD = 16;
            ColumnaValidacionYtg100 = 17;
            ColumnaValidacionYtgPonderado = 18;
            ColumnaValidacionTc = 19;
            ColumnaValidacionDiferencia = 20;

            Hoja = HojaVariacion;
        }

        public Variacion(Oportunidad oportunidad, Oportunidad oportunidadAnterior): this()
        {
            
            Cuenta = oportunidad.Cuenta;
            Nombre = oportunidad.Nombre;
            Codigo = oportunidad.Codigo;
            Responsable = oportunidad.Responsable;

            Fase = oportunidad.Fase;
            FaseAnterior = oportunidadAnterior.Fase;

            Ponderado = oportunidad.Ponderado;
            PonderadoAnterior = oportunidadAnterior.Ponderado;

            ImporteUSD(oportunidad.ImporteUSD());
            ImporteUSDAnterior = oportunidadAnterior.ImporteUSD();

            Monto(oportunidad.Monto());
            MontoAnterior = oportunidadAnterior.Monto();

            Probabilidad(oportunidad.Probabilidad());
            ProbabilidadAnterior = oportunidadAnterior.Probabilidad();
            
            HojaActual = oportunidad.Hoja;
            HojaAnterior = oportunidadAnterior.Hoja;

        }

        public override void CargarDatos(ExcelWorksheet hoja, int i)
        {
            throw new NotImplementedException();
        }

        public void PegarDatos(ExcelWorksheet hojaVariaciones, int filaVariacion)
        {
            hojaVariaciones.Cells[filaVariacion, ColumnaCuenta].Value = Cuenta;
            hojaVariaciones.Cells[filaVariacion, ColumnaCodigo].Value = Codigo;
            hojaVariaciones.Cells[filaVariacion, ColumnaOportunidad].Value = Nombre;
            hojaVariaciones.Cells[filaVariacion, ColumnaResponsable].Value = Responsable;
            hojaVariaciones.Cells[filaVariacion, ColumnaFase].Value = Fase;
            hojaVariaciones.Cells[filaVariacion, ColumnaFaseAnterior].Value = FaseAnterior;
            hojaVariaciones.Cells[filaVariacion, ColumnaImporteUSD].Value = ImporteUSD();
            hojaVariaciones.Cells[filaVariacion, ColumnaImporteUSDAnterior].Value = ImporteUSDAnterior;
            hojaVariaciones.Cells[filaVariacion, ColumnaMonto].Value = Monto();
            hojaVariaciones.Cells[filaVariacion, ColumnaMontoAnterior].Value = MontoAnterior;
            hojaVariaciones.Cells[filaVariacion, ColumnaPonderado].Value = Ponderado;
            hojaVariaciones.Cells[filaVariacion, ColumnaPonderadoAnterior].Value = PonderadoAnterior;
            hojaVariaciones.Cells[filaVariacion, ColumnaProbabilidad].Value = Probabilidad();
            hojaVariaciones.Cells[filaVariacion, ColumnaProbabilidadAnterior].Value = ProbabilidadAnterior;
            
            hojaVariaciones.Cells[filaVariacion, ColumnaValidacionTc].Value = 0;
            hojaVariaciones.Cells[filaVariacion, ColumnaValidacionYTD].Value = CalcularValidacionYTD();
            hojaVariaciones.Cells[filaVariacion, ColumnaValidacionYtg100].Value = CalcularValidacionYTG100();
            hojaVariaciones.Cells[filaVariacion, ColumnaValidacionYtgPonderado].Value = CalcularValidacionYTGPonderado();
            hojaVariaciones.Cells[filaVariacion, ColumnaValidacionDiferencia].Value = Ponderado - PonderadoAnterior;
        }

        public double CalcularValidacionYTD()
        {
            var acumulado = 0.0;
            if(HojaActual == Oportunidad.HojaYTD) acumulado = Ponderado;
            if(HojaAnterior == Oportunidad.HojaYTD) acumulado = acumulado - PonderadoAnterior;

            return acumulado;
        }

        public double CalcularValidacionYTG100()
        {
            var acumulado = 0.0;
            if (HojaActual == Oportunidad.HojaYTG100) acumulado = Ponderado;
            if (HojaAnterior == Oportunidad.HojaYTG100) acumulado = acumulado - PonderadoAnterior;

            return acumulado;
        }

        public double CalcularValidacionYTGPonderado()
        {
            var acumulado = 0.0;
            if (HojaActual == Oportunidad.HojaYTGPonderado) acumulado = Ponderado;
            if (HojaAnterior == Oportunidad.HojaYTGPonderado) acumulado = acumulado - PonderadoAnterior;

            return acumulado;
        }

        public bool SigueIgual()
        {
            if( HojaActual == HojaAnterior && Ponderado == PonderadoAnterior && ImporteUSD() == ImporteUSDAnterior 
                && Monto() == MontoAnterior && Probabilidad() == ProbabilidadAnterior)
            {
                return true;
            }

            return false;
        }
    }
}
