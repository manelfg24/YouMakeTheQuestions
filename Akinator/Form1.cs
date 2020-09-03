using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

namespace Akinator
{
    public partial class Form1 : Form
    {
        public static System.Windows.Forms.TextBox tbEnviar;
        public static System.Windows.Forms.TextBox tbPreguntaGuess;
        public static System.Windows.Forms.Button buttonSi;
        public static System.Windows.Forms.Button buttonNo;
        public static double coincidencia;
        public static string preguntaGuess;
        public static string excelBaseFilePath = @"C:\Users\manel\source\repos\YouMakeTheQuestions\excels\preguntasBase.xlsx";
        public static Excel.Application excelApp;
        public static Excel.Workbook activeWorkbook;
        public static Excel.Sheets activeSheets;
        public static Excel.Worksheet workingSheet;
        public static string sinonimosUrl = @"http://sesat.fdi.ucm.es:8080/servicios/rest/sinonimos/json/";
        public static HttpClient client;
        public static List<List<string>> sinonimosFrase;

        public Form1()
        {
            InitializeComponent();
            tbEnviar = textBox1;
            tbPreguntaGuess = textBox2;
            buttonSi = button2;
            buttonNo = button3;

            buttonSi.Enabled = false;
            buttonNo.Enabled = false;

            //Excel initialization
            //TODO no olvidar cerrar y liberar lo que toque de excel
            SetupExcel();

            client = null;

        }

        private static void SetupExcel()
        {
            excelApp = new Excel.Application();
            activeWorkbook = excelApp.Workbooks.Open(excelBaseFilePath);
            activeSheets = activeWorkbook.Worksheets;
            workingSheet = (Worksheet)activeSheets[1];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string preguntaNew = tbEnviar.Text;

            List<string> palabrasNew = new List<string>(preguntaNew.Split(' '));

            List<string> conjPreguntasBase = GetPreguntasExcel();

            coincidencia = 0.0;

            sinonimosFrase = new List<List<string>>();
            for(int i= 0; i < palabrasNew.Count; i++)
            {
                List<string> sinonimosPalabra = GetSinonimos(sinonimosUrl + palabrasNew[i]);
                if (sinonimosPalabra.Count > 0)
                {
                    sinonimosFrase.Add(sinonimosPalabra);
                }
            }

            bool found = false;


            foreach(string preguntaBase in conjPreguntasBase) //para cada pregunta de nuestra BD
            {
                
                if (preguntaBase.Equals(preguntaNew))
                {
                    tbPreguntaGuess.Text = "Ya en BD! - " + preguntaBase;
                    found = true;
                    buttonSi.Enabled = false;
                    buttonNo.Enabled = false;
                    break;
                }
                
                double count = 0.0;
                List<string> palabrasBase = new List<string>(preguntaBase.Split(' '));
                foreach (string palabraBase in palabrasBase) //para cada palabra de cada pregunta de nuestra BD
                {
                    foreach (string palabraNew in palabrasNew) //para cada palabra de la pregunta nueva
                    {
                        //TODO anadir comparacion sinonimos cuando este arreglado
                        if (palabraBase.Equals(palabraNew)) count = count + 1.0; //Si las palabras son iguales o sinonimas, + coincidencia
                    }

                    if (IsInMatrix(sinonimosFrase, palabraBase)) count = count + 0.8;
                }

                double ratio = (double)count / Math.Max(palabrasBase.Count, palabrasNew.Count);
                if (ratio > coincidencia)
                {
                    coincidencia = ratio;
                    preguntaGuess = preguntaBase;
                }
                
            }

            //TODO quitar chivatillo de coincidencia cuando toque
            if (!found)
            {
                buttonSi.Enabled = true;
                buttonNo.Enabled = true;
                tbPreguntaGuess.Text = preguntaGuess + " - coincidencia: " + coincidencia * 100 + "%";
            }
            


        }

        public bool IsInMatrix(List<List<string>> matrix, string palabra)
        {
            bool found = false;
            int matWidth = matrix.Count;
            string aux;
            int i = 0;
            while (!found && i < matWidth)
            {
                int j = 0;
                while (!found && j < matrix[i].Count)
                {
                    aux = matrix[i][j];
                    aux = Regex.Replace(aux.Normalize(NormalizationForm.FormD), @"[^a-zA-z0-9 ]+", "");
                    if (aux.Equals(palabra)) found = true;
                    j++;
                }
                i++;
            }
            return found;
        }

        public List<string> DeserializeJsonSinonimos(string json)
        {
            List<string> list = new List<string>();

            try
            {
                JsonSinonimos sinonimosList = JsonConvert.DeserializeObject<JsonSinonimos>(json);
                foreach (Sinonimo sinonimo in sinonimosList.sinonimos)
                {
                    list.Add(sinonimo.sinonimo);
                }
            }
            catch
            {

            }

            return list;
        }

        public List<string> GetSinonimos(string path)
        {
            if (client == null)
            {
                HttpClientHandler handler = new HttpClientHandler()
                {
                    AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
                };
                client = new HttpClient(handler);
            }
            HttpResponseMessage response = client.GetAsync(path).Result;
            response.EnsureSuccessStatusCode();
            string result = response.Content.ReadAsStringAsync().Result;

            List<string> sinonimos = DeserializeJsonSinonimos(result);
            return sinonimos;
        }


        //Funcion que  abre el documento Excel y devuelve el contenido
        //de la primera columna como List.
        public List<string> GetPreguntasExcel()
        {
            List<string> preguntasList = new List<string>();

            Range range = workingSheet.UsedRange;
            int rowCount = range.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                if (range.Cells[i, 2] != null)
                {
                    preguntasList.Add(range.Cells[i, 2].Value.ToString());
                }
            }

            return preguntasList;
        }


        //Funcion que abre un documento Excel y añade
        //a la primera columna la pregunta proporcionada.
        public void AddPreguntaIneditaExcel(string pregunta)
        {
            Range range = workingSheet.UsedRange;
            int rowCount = range.Rows.Count;

            range.Cells[rowCount + 1, 1] = "2";
            range.Cells[rowCount + 1, 2] = pregunta;

            activeWorkbook.Save();
        }

        //Funcion que, dado un filepath valido, abre un documento Excel y añade
        //a la primera columna la pregunta proporcionada.
        public void AddPreguntaSemejanteExcel(string preguntaNueva, string preguntaBase)
        {
            //hacer cosas
            Range range = workingSheet.UsedRange;
            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;

            //buscar la cell que coincida con preguntaBase
            int i = 1;
            int j = 2;
            bool found = false;
            while (!found && (j <= colCount))
            {
                while(!found && (i <= rowCount))
                {
                    if (range.Cells[i,j] != null)
                    {
                        if (range.Cells[i,j].Value.ToString().Equals(preguntaBase))
                        {
                            found = true;
                            int nextPos = Convert.ToInt32(range.Cells[i, 1].Value2.ToString()) + 1; //+1 para ponerme en la proxima posicion
                            range.Cells[i, nextPos] = preguntaNueva;
                            range.Cells[i, 1] = Convert.ToString(nextPos);
                        }
                    }
                    i++;
                }
                j++;
            }

            activeWorkbook.Save();
        }


        //Boton SI
        private void button2_Click(object sender, EventArgs e)
        {
            string preguntaNew = tbEnviar.Text;
            string preguntaBase = preguntaGuess;

            AddPreguntaSemejanteExcel(preguntaNew, preguntaBase);

            ResetForm();
        }

        //Boton NO
        private void button3_Click(object sender, EventArgs e)
        {
            string preguntaNew = tbEnviar.Text;

            AddPreguntaIneditaExcel(preguntaNew);

            ResetForm();
        }

        //Boton SALIR
        private void button4_Click(object sender, EventArgs e)
        {
            activeWorkbook.Close();
            excelApp.Quit();
            this.Dispose();
            this.Close(); 
        }

        public void ResetForm()
        {
            tbEnviar.Text = "";
            tbPreguntaGuess.Text = "";
        }
    }
}
