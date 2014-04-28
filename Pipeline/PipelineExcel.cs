using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Pipeline
{
    public static class PipelineExcel
    {
        public static void CrearVariacion(string pathExcelAnterior, string pathExcelActual, ProgressBar progressBar)
        {

            //Anterior
            var excelAnterior = new ListaHojasExcel(pathExcelAnterior);
            progressBar.Value ++;

            //Actual
            var excelActual = new ListaHojasExcel(pathExcelActual);
            progressBar.Value++;
    
            var hojaVariaciones = excelActual.Excel.Workbook.Worksheets[Oportunidad.HojaVariacion];
            var filaVariacion = 5;
            //var variacionHeader = new Variacion();
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaCuenta].Value = "Cuenta";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaCodigo].Value = "Codigo";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaOportunidad].Value = "Nombre Oportunidad";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaResponsable].Value = "Responsable";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaFase].Value = "Fase";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaFaseAnterior].Value = "FaseAnterior";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaImporteUSD].Value = "ImporteUSD";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaImporteUSDAnterior].Value = "ImporteUSDAnterior";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaMonto].Value = "Monto";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaMontoAnterior].Value = "MontoAnterior";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaPonderado].Value = "Ponderado";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaPonderadoAnterior].Value = "PonderadoAnterior";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaProbabilidad].Value = "Probabilidad";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaProbabilidadAnterior].Value = "ProbabilidadAnterior";
            ////Separacion
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionTc].Value = "TC";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionYTD].Value = "YTD";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionYtg100].Value = "YTG100";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionYtgPonderado].Value = "YTGPonderado";
            //hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionDiferencia].Value = "Diferencia";
            filaVariacion++;

            excelActual.EliminarIguales(excelAnterior);
            progressBar.Value++;
            //excelActual.EliminarAgrupados(excelAnterior);
            //progressBar.Step++;
            var listaVariacionesIguales = excelActual.DiferenciaEntreIguales(excelAnterior);
            progressBar.Value++;

            var listaVariacionesNuevas = excelActual.DiferenciaAntesNoExistianEnElAnterior(excelAnterior);
            progressBar.Value++;

            var listaVariacionesAnteriores = excelActual.DiferenciaAntesNoExistenEnElNuevo(excelAnterior);
            progressBar.Value++;

            

            var listaVariacionesTotales = listaVariacionesIguales.Union(listaVariacionesNuevas).Union(listaVariacionesAnteriores);
            var variaciones = listaVariacionesTotales.ToList();
            progressBar.Value++;

            foreach (var variacion in variaciones)
            {
                variacion.PegarDatos(hojaVariaciones, filaVariacion);
                filaVariacion++;
            }

            progressBar.Value++;
            excelActual.Excel.Save();
            excelActual.Excel.Dispose();

        }

        public static ExcelWorksheet CreateSheet(ExcelPackage package, string sheetName)
        {
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet != null)
            {
                package.Workbook.Worksheets.Delete(sheetName);
            }
            package.Workbook.Worksheets.Add(sheetName);
            worksheet = package.Workbook.Worksheets[sheetName];
            worksheet.Name = sheetName; //Setting Sheet's name
            worksheet.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            worksheet.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

            return worksheet;
        }
    }
}
