using System;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;


namespace Pipeline
{
    public partial class Main : Form
    {
        private string _pathExcelAnterior="";
        private string _pathExcelActual="";
        private string _pathExcelNuevo = "";

        public Main()
        {
            InitializeComponent();
            this.progressBar.Hide();
        }

        private void btnAbrirAnterior_Click(object sender, EventArgs e)
        {

            var fileDialog = new OpenFileDialog();
            fileDialog.Title = "Abrir archivo Excel";
            fileDialog.Filter = "Excel file|*.xlsx";
            //fileDialog.InitialDirectory = @"C:\";
            var dialogResult =fileDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                _pathExcelAnterior = fileDialog.FileName;
                txtExcelAnterior.Text = _pathExcelAnterior;
            }

        }

        private void btnAbrirActual_Click(object sender, EventArgs e)
        {

            var fileDialog = new OpenFileDialog();
            fileDialog.Title = "Abrir archivo Excel";
            fileDialog.Filter = "Excel file|*.xlsx";
            var dialogResult = fileDialog.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                _pathExcelActual = fileDialog.FileName;
                txtExcelActual.Text = _pathExcelActual;
            }

        }

        private void btnEjecutar_Click(object sender, EventArgs e)
        {

            if (_pathExcelAnterior.Trim() == "")
            {
                MessageBox.Show("Debe seleccionar un excel anterior");
                return;
            }
            
            if (_pathExcelActual.Trim() == "")
            {
                MessageBox.Show("Debe seleccionar un excel actual");
                return;
            }

            //_pathExcelNuevo = "";
            //var saveFileDialog = new SaveFileDialog();
            //saveFileDialog.Title = "Crear archivo Excel";
            //saveFileDialog.Filter = "Excel file|*.xlsx";
            //saveFileDialog.FileName = "Variacion";
            //var dialogResult = saveFileDialog.ShowDialog();

            //if (dialogResult != DialogResult.OK)
            //{
            //    return;
            //}
            //_pathExcelNuevo = saveFileDialog.FileName;

            //if (_pathExcelNuevo == "") return;

            this.btnEjecutar.Hide();

            this.progressBar.Maximum = 9;
            this.progressBar.Minimum = 1;
            this.progressBar.Show();
            this.progressBar.Value = 1;

            try
            {
                PipelineExcel.CrearVariacion(_pathExcelAnterior, _pathExcelActual, this.progressBar);
                MessageBox.Show("Ejecucion exitosa");
            }
            catch(IOException ioExeption)
            {
                MessageBox.Show("El archivo de excel se encuentra Actualmente abierto: " + ioExeption.Message);
            }
            this.progressBar.Hide();
            this.btnEjecutar.Show();
            
        }
    }
}
