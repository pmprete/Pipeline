﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Pipeline
{
    public class YTG100: Oportunidad
    {
        public YTG100()
        {
            ColumnaCuenta = 2;
            ColumnaOportunidad = 3;
            ColumnaCodigo = 4;
            ColumnaResponsable = 5;
            ColumnaFase = 6;
            ColumnaProbabilidad = 7;
            ColumnaImporteUSD = 8;
            ColumnaPonderado = 9;
            ColumnaFechaDeIngreso = 10;

            Hoja = HojaYTG100;

        }

        public override void CargarDatos(ExcelWorksheet hoja, int i)
        {
            Cuenta = hoja.GetValue<string>(i, ColumnaCuenta);
            Nombre = hoja.GetValue<string>(i, ColumnaOportunidad);
            Codigo = hoja.GetValue<int>(i, ColumnaCodigo);
            Responsable = hoja.GetValue<string>(i, ColumnaResponsable);
            Fase = hoja.GetValue<string>(i, ColumnaFase);
            Probabilidad = hoja.GetValue<double>(i, ColumnaProbabilidad);
            ImporteUSD = Math.Round(hoja.GetValue<double>(i, ColumnaImporteUSD));
            Ponderado = Math.Round(hoja.GetValue<double>(i, ColumnaPonderado));
            FechaDeIngreso = ConvertirExcelAFecha(hoja, i, ColumnaFechaDeIngreso);
            Monto = Math.Round(Ponderado);
        }
    }
}
