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

namespace Akinator
{
    public partial class Form1 : Form
    {
        public static TextBox tbEnviar;
        public static TextBox tbPreguntaGuess;
        public static double coincidencia;
        public static string preguntaGuess;
        public Form1()
        {
            InitializeComponent();
            tbEnviar = textBox1;
            tbPreguntaGuess = textBox2;


           

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
                        if (palabraBase.Equals(palabraNew)) ++count; //Si las palabras son iguales o sinonimas, + coincidencia
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

        public ArrayList GetSinonimos(string palabra)
        //Funcion que, dada una palabra, devuelve un ArrayList de strings con los sinonimos
        //en español de esta palabra.
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
    }
}
