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

namespace Akinator
{
    public partial class Form1 : Form
    {
        public static System.Windows.Forms.TextBox tbEnviar;
        public static System.Windows.Forms.TextBox tbPreguntaGuess;
        public static double coincidencia;
        public static string preguntaGuess;
        public static string excelBaseFilePath = @"C:\Users\manel\source\repos\YouMakeTheQuestions\excels\preguntasBase.xlsx";
        public static Excel.Application excelApp;
        public static Excel.Workbook activeWorkbook;
        public static Excel.Sheets activeSheets;
        public static Excel.Worksheet workingSheet;

        public Form1()
        {
            InitializeComponent();
            tbEnviar = textBox1;
            tbPreguntaGuess = textBox2;

            //Excel initialization
            //TODO no olvidar cerrar y liberar lo que toque de excel
            SetupExcel();

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
            //TODO FUNCION TOCHA
            
            string preguntaNew = tbEnviar.Text;

            ArrayList palabrasNew = new ArrayList(preguntaNew.Split(' '));

            ArrayList conjPreguntasBase = new ArrayList();
            conjPreguntasBase = GetPreguntasExcel();

            coincidencia = 0.0;


            foreach(string preguntaBase in conjPreguntasBase) //para cada pregunta de nuestra BD
            {
                int count = 0;
                ArrayList palabrasBase = new ArrayList(preguntaBase.Split(' '));
                foreach (string palabraBase in palabrasBase) //para cada palabra de cada pregunta de nuestra BD
                {
                    foreach (string palabraNew in palabrasNew) //para cada palabra de la pregunta nueva
                    {
                        //TODO anadir comparacion sinonimos cuando este arreglado
                        if (palabraBase.Equals(palabraNew) || AreSinonimas(palabraBase,palabraNew)) ++count; //Si las palabras son iguales o sinonimas, + coincidencia
                    }
                }

                double ratio = (double)count / Math.Max(palabrasBase.Count, palabrasNew.Count);
                if (ratio > coincidencia)
                {
                    coincidencia = ratio;
                    preguntaGuess = preguntaBase;
                }
                
            }

            //TODO quitar chivatillo de coincidencia cuando toque
            tbPreguntaGuess.Text = preguntaGuess + " - coincidencia: " + coincidencia*100 + "%";



        }

        public bool AreSinonimas(string palabraA, string palabraB)
        {
            bool comp = GetSinonimos(palabraA).Contains(palabraB);
            return comp;
        }

        //Funcion que, dada una palabra, devuelve un ArrayList de strings con los sinonimos
        //en español de esta palabra.
        public ArrayList GetSinonimos(string palabra)
        {
            ArrayList sinonimosArray = new ArrayList();
            var appWord = new Microsoft.Office.Interop.Word.Application();
            object objLanguage = Microsoft.Office.Interop.Word.WdLanguageID.wdSpanish;
            Microsoft.Office.Interop.Word.SynonymInfo si = appWord.get_SynonymInfo(palabra, ref (objLanguage));
            foreach (var meaning in (si.MeaningList as Array))
            {
                sinonimosArray.Add(meaning.ToString());
            }
            appWord.Quit(); //include this to ensure the related process (winword.exe) is correctly closed. 
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appWord);
            objLanguage = null;
            appWord = null;
            return sinonimosArray;
        }

        
        //Funcion que  abre el documento Excel y devuelve el contenido
        //de la primera columna como ArrayList.
        public ArrayList GetPreguntasExcel()
        {
            ArrayList preguntasArray = new ArrayList();

            Range range = workingSheet.UsedRange;
            int rowCount = range.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                if (range.Cells[i, 1] != null)
                {
                    preguntasArray.Add(range.Cells[i, 1].Value.ToString());
                }
            }

            return preguntasArray;
        }


        //Funcion que abre un documento Excel y añade
        //a la primera columna la pregunta proporcionada.
        public void AddPreguntaIneditaExcel(string pregunta)
        {
            Range range = workingSheet.UsedRange;
            int rowCount = range.Rows.Count;

            range.Cells[rowCount + 1, 1] = pregunta;
        }

        //Funcion que, dado un filepath valido, abre un documento Excel y añade
        //a la primera columna la pregunta proporcionada.
        public void AddPreguntaSemejanteExcel(string preguntaNueva, string preguntaBase)
        {
            //TODO hacer cosas
            Range range = workingSheet.UsedRange;
            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;

            //TODO buscar la cell que coincida con preguntaBase
            int i = 1;
            int j = 1;
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
                            while (range.Cells[i, j] != null) j++;
                            range.Cells[i, j] = preguntaNueva;
                        }
                    }
                    i++;
                }
                j++;
            }
        }
    }
}
