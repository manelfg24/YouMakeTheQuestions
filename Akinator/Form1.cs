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
        public Form1()
        {
            InitializeComponent();
            tbEnviar = textBox1;
            tbPreguntaGuess = textBox2;

            //olalalaa
           

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //TODO FUNCION TOCHA
            
            string preguntaNew = tbEnviar.Text;

            ArrayList palabrasNew = new ArrayList(preguntaNew.Split(' '));

            ArrayList conjPreguntasBase = new ArrayList(); //TODO conseguir coger de excel o calc todas las preguntas
            conjPreguntasBase.Add("Es guapo");
            conjPreguntasBase.Add("Vive en el mar");
            conjPreguntasBase.Add("Tiene el pelo rubio");
            conjPreguntasBase.Add("Es real");

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
            return GetSinonimos(palabraA).Contains(palabraB);
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

        
        //Funcion que, dado un filepath valido, abre un documento Excel y devuelve el contenido
        //de la primera columna como ArrayList.
        public ArrayList GetPreguntasExcel(string filepath)
        {
            ArrayList preguntasArray = new ArrayList();

            var appExcel = new Microsoft.Office.Interop.Excel.Application();
            var workbook = appExcel.Workbooks.Open(filepath);
            var sheets = workbook.Worksheets;
            var workSheet = (Worksheet)sheets[1];

            Range range = workSheet.UsedRange;
            int rowCount = range.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                if (range.Cells[i, 1] != null)
                {
                    preguntasArray.Add(range.Cells[i, 1].Value.ToString());
                }
            }

            workbook.Close();
            appExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
            workbook = null;
            appExcel = null;

            return preguntasArray;

        }


        //Funcion que, dado un filepath valido, abre un documento Excel y añade
        //a la primera columna la pregunta proporcionada.
        public void AddPreguntaExcel(string filepath, string pregunta)
        {
            ArrayList preguntasArray = new ArrayList();

            var appExcel = new Microsoft.Office.Interop.Excel.Application();
            var workbook = appExcel.Workbooks.Open(filepath);
            var sheets = workbook.Worksheets;
            var workSheet = (Worksheet)sheets[1];

            //TODO hacer cosas
            Range range = workSheet.UsedRange;
            int rowCount = range.Rows.Count;

            range.Cells[rowCount + 1, 1] = pregunta;

            workbook.Close();
            appExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
            workbook = null;
            appExcel = null;

        }
    }
}
