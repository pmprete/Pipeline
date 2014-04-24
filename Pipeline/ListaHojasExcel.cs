using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Pipeline
{
    public class ListaHojasExcel
    {
        public Dictionary<int, List<Oportunidad>> Hojas { get; set; }


        public ListaHojasExcel()
        {
            Hojas = new Dictionary<int, List<Oportunidad>>
                        {
                            {Oportunidad.HojaYTD, new List<Oportunidad>()},
                            {Oportunidad.HojaYTG100, new List<Oportunidad>()},
                            {Oportunidad.HojaYTGPonderado, new List<Oportunidad>()},
                            {Oportunidad.HojaPerdidas, new List<Oportunidad>()}
                        };
        }

        public ListaHojasExcel(string pathExcel) : this()
        {
            var archivo = new FileInfo(pathExcel);
            var excel = new ExcelPackage(archivo);

            foreach (var hoja in Hojas)
            {
                var indice = hoja.Key;
                var hojaActual = excel.Workbook.Worksheets[indice];
                var filaActual = 5;
                while (!String.IsNullOrEmpty(hojaActual.GetValue<string>(filaActual, 2)))
                {

                    var oportunidad = Oportunidad.CrearOportunidad(indice);
                    oportunidad.CargarDatos(hojaActual, filaActual);
                    hoja.Value.Add(oportunidad);
                    filaActual++;
                }
            }
            
            excel.Dispose();
        }

        public List<Variacion> ObtenerVariaciones(ListaHojasExcel excelAnterior)
        {
            EliminarIguales(excelAnterior);
            EliminarAgrupados(excelAnterior);

            var listaVariaciones = new List<Variacion>();
            foreach (var hoja in Hojas)
            {
                foreach (var hojaAnterior in excelAnterior.Hojas)
                {
                    foreach (var oportunidad in hoja.Value.ToList())
                    {
                        var oportunidadAnterior = hojaAnterior.Value.FirstOrDefault(x=> x.Codigo == oportunidad.Codigo && x.FechaDeIngreso == oportunidad.FechaDeIngreso);

                        if (oportunidadAnterior != null)
                        {
                            var variacion = new Variacion(oportunidad, oportunidadAnterior);
                            listaVariaciones.Add(variacion);
                            hojaAnterior.Value.Remove(oportunidadAnterior);
                            hoja.Value.Remove(oportunidad);
                        }
                    }
                }
            }

            return listaVariaciones;
        }

         public void EliminarIguales(ListaHojasExcel excelAnterior)
        {
            foreach (var hoja in Hojas)
            {
                var nterior = excelAnterior.Hojas[hoja.Key];
                var iguales = hoja.Value.Where(x => nterior.Any(y => x.Iguales(y))).ToList();
                hoja.Value.RemoveAll(x => iguales.Any(x.Iguales));
                nterior.RemoveAll(x => iguales.Any(x.Iguales));
            }
            
        }

        public void EliminarAgrupados(ListaHojasExcel excelAnterior)
        {
            foreach (var hoja in Hojas)
            {
                var codigosActuales = from oportunidad in hoja.Value
                              let grupo = new
                              {
                                  Codigo = oportunidad.Codigo,
                                  Fase = oportunidad.Fase,
                                  Probabilidad = oportunidad.Probabilidad,
                                  FechaDeIngreso = oportunidad.FechaDeIngreso
                              }
                              group oportunidad by grupo into t
                              select new
                              {
                                  Codigo = t.Key.Codigo,
                                  Fase = t.Key.Fase,
                                  Probabilidad = t.Key.Probabilidad,
                                  FechaDeIngreso = t.Key.FechaDeIngreso,
                                  Total = t.Sum(oportunidad => oportunidad.Ponderado)
                              };

                var codigosAnteriores = from oportunidad in excelAnterior.Hojas[hoja.Key]
                                      let grupo = new
                                      {
                                          Codigo = oportunidad.Codigo,
                                          Fase = oportunidad.Fase,
                                          Probabilidad = oportunidad.Probabilidad,
                                          FechaDeIngreso = oportunidad.FechaDeIngreso
                                      }
                                      group oportunidad by grupo into t
                                      select new
                                      {
                                          Codigo = t.Key.Codigo,
                                          Fase = t.Key.Fase,
                                          Probabilidad = t.Key.Probabilidad,
                                          FechaDeIngreso = t.Key.FechaDeIngreso,
                                          Total = t.Sum(oportunidad => oportunidad.Ponderado)
                                      };

                var gruposIguales = codigosActuales.Where(x => codigosAnteriores.Any(o => x.Codigo == o.Codigo && x.Fase == o.Fase 
                    && x.Probabilidad == o.Probabilidad && x.FechaDeIngreso == o.FechaDeIngreso && x.Total == o.Total)).ToList();

                hoja.Value.RemoveAll(x => gruposIguales.Any(o => x.Codigo == o.Codigo && x.Fase == o.Fase
                    && x.Probabilidad == o.Probabilidad && x.FechaDeIngreso == o.FechaDeIngreso));

                excelAnterior.Hojas[hoja.Key].RemoveAll(x => gruposIguales.Any(o => x.Codigo == o.Codigo && x.Fase == o.Fase
                    && x.Probabilidad == o.Probabilidad && x.FechaDeIngreso == o.FechaDeIngreso));


            }
            
        }

    }
}
