﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Pipeline
{
    public class YTGPonderado: Oportunidad
    {

        public YTGPonderado()
        {
            ColumnaCuenta = 2;
            ColumnaOportunidad = 3;
            ColumnaCodigo = 4;
            ColumnaResponsable = 5;
            ColumnaFase = 6;
            ColumnaProbabilidad = 7;
            ColumnaImporteUSD = 8;
            ColumnaMonto = 9;
            ColumnaPonderado = 10;
            ColumnaFechaDeIngreso = 11;

            Hoja = HojaYTGPonderado;
        }

        public override void CargarDatos(ExcelWorksheet hoja, int i)
        {
            Cuenta = hoja.GetValue<string>(i, ColumnaCuenta);
            Nombre = hoja.GetValue<string>(i, ColumnaOportunidad);
            Codigo = hoja.GetValue<int>(i, ColumnaCodigo);
            Responsable = hoja.GetValue<string>(i, ColumnaResponsable);
            Fase = hoja.GetValue<string>(i, ColumnaFase);
            Probabilidad = hoja.GetValue<double>(i, ColumnaProbabilidad);
            ImporteUSD =hoja.GetValue<double>(i, ColumnaImporteUSD);
            Monto = hoja.GetValue<double>(i, ColumnaMonto);
            Ponderado = hoja.GetValue<double>(i, ColumnaPonderado);
            FechaDeIngreso = hoja.GetValue<string>(i, ColumnaFechaDeIngreso);
        }
    }
}
