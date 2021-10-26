using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data.SqlClient;
using DataTable = System.Data.DataTable;
using System.IO;

using System.Runtime.InteropServices;

namespace ProjetoTicunas
{
    public partial class Form1 : Form
    {

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]

        private static extern IntPtr CreateRoundRectRgn(
        int nLeftRect,
        int nTopRect,
        int nRightRect,
        int nBottomRect,
        int nWidthEllipse,
        int nHeightEllipse
            );


        List<string> lista = new List<string>();

        string pathfile = @"C:\Users\arthu\OneDrive\Área de Trabalho\Curso C#\Projetos\Ticunas\EstoqueTicunas2.xlsx";
        string sheet = "Peças superiores";

        public Form1()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            this.Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            PopulateComboBoxes();
            InicializaSistema();
        }

        //Botão de filtro
        private void Filtrar_Click(object sender, EventArgs e)
        {
            string CombName;

            //Passa por todos os comboboxes no form
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is ComboBox)
                {
                    if (ctrl.Text != "-")
                    {
                        CombName = ctrl.Name;
                        GetInfo(CombName);                       
                    }
                }

            }

            AplicaFiltros();

            lista.Clear();        
        }

        //Desfiltra tudo do estoque
        private void Unfilter_Click_1(object sender, EventArgs e)
        {
            ImportExcelToDt();
        }

        //Popula DataGridView com source de Excel
        public void ImportExcelToDt()
        {
            #region Importa Excel para DataGridView
            //Importa excel para DataGridView 


            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            pathfile +
                            ";Extended Properties='Excel 8.0;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + sheet + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            dataGridView1.DataSource = data;
            #endregion

        }

        //Inicializa sistema e traz informações
        public void InicializaSistema()
        {
            //string Filename = "EstoqueTicunas2.xlsx";
            #region Excel
            Boolean Opened = IsOpened(pathfile);

            //Verifica se excel está aberto
            if (Opened)
            {
                OpenFile();
                //Importa conteudo para DataGridView
                ImportExcelToDt();
            }
            else //Se não está aberto, abre excel
            {
                ImportExcelToDt();
            }
            #endregion
        }

        //Popula Comboboxes de filtros
        public void PopulateComboBoxes()
        {
            #region Comboboxes
            //Popula ComboBox de Generos
            var dataSourceGenero = new List<Generos>();
            dataSourceGenero.Add(new Generos() { Genero = "-" });
            dataSourceGenero.Add(new Generos() { Genero = "Masculino" });
            dataSourceGenero.Add(new Generos() { Genero = "Feminino" });


            this.FiltroGenero.DataSource = dataSourceGenero;
            this.FiltroGenero.DisplayMember = "Genero";
            this.FiltroGenero.ValueMember = "Genero";

            //Popula ComboBox de Locais            
            var dataSourceLocal = new List<Localizações>();
            dataSourceLocal.Add(new Localizações() { Local = "-" });
            dataSourceLocal.Add(new Localizações() { Local = "Loja 1" });
            dataSourceLocal.Add(new Localizações() { Local = "Loja 2" });


            this.CbLocal.DataSource = dataSourceLocal;
            this.CbLocal.DisplayMember = "Local";
            this.CbLocal.ValueMember = "Local";

            //Popula ComboBox de Tamanhos
            var dataSourceQuantity = new List<Tamanho>();
            dataSourceQuantity.Add(new Tamanho() { Tam = "-" });
            dataSourceQuantity.Add(new Tamanho() { Tam = "PP" });
            dataSourceQuantity.Add(new Tamanho() { Tam = "P" });
            dataSourceQuantity.Add(new Tamanho() { Tam = "M" });
            dataSourceQuantity.Add(new Tamanho() { Tam = "G" });
            dataSourceQuantity.Add(new Tamanho() { Tam = "GG" });


            this.cbTamanho.DataSource = dataSourceQuantity;
            this.cbTamanho.DisplayMember = "Tam";
            this.cbTamanho.ValueMember = "Tam";
            #endregion

        }

        //Classes de comboboxes
        #region classes
        public class Generos
        {
            public string Genero { get; set; }
        }

        public class Localizações
        {
            public string Local { get; set; }
        }

        public class Tamanho
        {
            public string Tam { get; set; }
        }
        #endregion

        //Verifica se planilha está aberta
        static bool IsOpened(string wbook)
        {
            try
            {
                File.Open(wbook, FileMode.Open, FileAccess.ReadWrite);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void OpenFile()
        {
            Excel excel = new Excel(pathfile, 1);
        }

        private void roundButton1_Click(object sender, EventArgs e)
        {

        }

        //Filtros
        public void GetInfo(string cbName)
        {
           
            if (cbName == "FiltroGenero")
            {
                //Adiciona valores a lista
                string columnName = "Genero";
                string filterValue = FiltroGenero.Text;

                lista.Add(columnName);
                lista.Add(filterValue);
            }
            else if (cbName == "cbTamanho")
            {

                //Adiciona valores a lista
                string TamColumn = "Tamanho";
                string TamFilter = cbTamanho.Text;

                lista.Add(TamColumn);
                lista.Add(TamFilter);

            }
            else if (cbName == "CbLocal")
            {
                //Adiciona valores a lista
                string LocalColumn = "Localização";
                string LocalFilter = CbLocal.Text;

                lista.Add(LocalColumn);
                lista.Add(LocalFilter);         
            }
        }

        public void AplicaFiltros()
        {
            int length = lista.Count;

            if(length > 0)
            {
                if (length == 2)
                {
                    string rowFilter2 = string.Format("[{0}] = '{1}'", lista[0], lista[1]);
                    (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = rowFilter2;
                }
                else if (length == 4)
                {
                    string rowFilter2 = string.Format("[{0}] = '{1}' AND [{2}] = '{3}'", lista[0], lista[1], lista[2], lista[3]);
                    (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = rowFilter2;
                }
                else
                {
                    string rowFilter2 = string.Format("[{0}] = '{1}' AND [{2}] = '{3}' AND [{4}] = '{5}'", lista[0], lista[1], lista[2], lista[3], lista[4], lista[5]);
                    (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = rowFilter2;
                }
            }  
        }

        public bool IsNullOrEmpty(string[] myStringArray)
        {
            if (myStringArray == null)
            {
                return true;
            }
            else
            {
                for (int i =0; i < myStringArray.Length; i++)
                {
                    if (myStringArray[0] == "" || myStringArray[0] == "0")
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            return false;    
        }
    }
}

