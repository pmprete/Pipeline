using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Pipeline
{
    public static class PipelineExcel
    {
        public static void CrearVariacion(string pathExcelAnterior, string pathExcelActual, string pathExcelNuevo)
        {
            //Anterior
            var excelAnterior = new ListaHojasExcel(pathExcelAnterior);
            
            //Actual
            var excelActual = new ListaHojasExcel(pathExcelActual);

            //Nuevo
            var archivoNuevo = new ExcelPackage(new FileInfo(pathExcelNuevo));
            
     
            var hojaVariaciones = CreateSheet(archivoNuevo, "Variacion");
            var filaVariacion = 1;
            var variacionHeader = new Variacion();
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaCuenta].Value = "Cuenta";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaCodigo].Value = "Codigo";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaOportunidad].Value = "Nombre Oportunidad";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaResponsable].Value = "Responsable";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaFase].Value = "Fase";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaFaseAnterior].Value = "FaseAnterior";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaImporteUSD].Value = "ImporteUSD";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaImporteUSDAnterior].Value = "ImporteUSDAnterior";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaMonto].Value = "Monto";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaMontoAnterior].Value = "MontoAnterior";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaPonderado].Value = "Ponderado";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaPonderadoAnterior].Value = "PonderadoAnterior";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaProbabilidad].Value = "Probabilidad";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaProbabilidadAnterior].Value = "ProbabilidadAnterior";
                                                 
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionTc].Value = "TC";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionYTD].Value = "YTD";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionYtg100].Value = "YTG100";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionYtgPonderado].Value = "YTGPonderado";
            hojaVariaciones.Cells[filaVariacion, variacionHeader.ColumnaValidacionDiferencia].Value = "Diferencia";
            filaVariacion++;

            var variaciones = excelActual.ObtenerVariaciones(excelAnterior);
            foreach (var variacion in variaciones)
            {
                variacion.PegarDatos(hojaVariaciones, filaVariacion);
                filaVariacion++;
            }

            archivoNuevo.Save();
            archivoNuevo.Dispose();

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
