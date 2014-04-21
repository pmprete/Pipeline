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
            var archivoAnterior = new FileInfo(pathExcelAnterior);
            var excelAnterior = new ExcelPackage(archivoAnterior);
            
            //Actual
            var archivoActual = new FileInfo(pathExcelActual);
            var excelActual = new ExcelPackage(archivoActual);

            //Nuevo
            var archivoNuevo = new FileInfo(pathExcelNuevo);
            var excelNuevo = new ExcelPackage(archivoNuevo);

            var listaHojas = new List<int>
                                 {
                                     Oportunidad.HojaYTD,
                                     Oportunidad.HojaYTG100,
                                     Oportunidad.HojaYTGPonderado,
                                     Oportunidad.HojaPerdidas
                                 };
     
            var hojaVariaciones = CreateSheet(excelNuevo, "Variacion");
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

            foreach (var indice in listaHojas)
            {
                var hojaActual = excelActual.Workbook.Worksheets[indice];
                var filaActual = 5;
                while (!String.IsNullOrEmpty(hojaActual.GetValue<string>(filaActual, 2)))
                {

                    var oportunidad = Oportunidad.CrearOportunidad(indice);
                    oportunidad.CargarDatos(hojaActual, filaActual);
                   
                    foreach (var indiceAnterior in listaHojas)
                    {
                        var hojaAnterior = excelAnterior.Workbook.Worksheets[indiceAnterior];
                        var filaAnterior = 5;
                        while (!String.IsNullOrEmpty(hojaAnterior.GetValue<string>(filaAnterior, 2)))
                        {
                            var oportunidadAnterior = Oportunidad.CrearOportunidad(indiceAnterior);
                            oportunidadAnterior.CargarDatos(hojaAnterior, filaAnterior);
                            if (oportunidad.Codigo == oportunidadAnterior.Codigo)
                            {
                                var variacion = new Variacion(oportunidad, oportunidadAnterior);
                                if (!variacion.SigueIgual())
                                {
                                    variacion.PegarDatos(hojaVariaciones, filaVariacion);
                                    filaVariacion++;
                                }
                            }

                            filaAnterior++;
                        }

                    }

                    filaActual++;
                }
            }
            


            excelNuevo.Save();
            excelAnterior.Dispose();
            excelActual.Dispose();
            excelNuevo.Dispose();

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
