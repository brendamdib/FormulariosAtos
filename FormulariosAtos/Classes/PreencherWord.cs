using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop;

namespace FormulariosAtos
{
    class PreencherWord
    {
        public void PreencherPorReplace(string CaminhoDocMatriz)
        {
            //Objeto a ser usado nos parâmetros opcionais
            object missing = System.Reflection.Missing.Value;

            //Abre a aplicação Word e faz cópia do objeto mapeado
            Word.Word.Application oApp = new Word.Word.Application();

            object template = CaminhoDocMatriz;

            Word.Word.Document oDoc = oApp.Documents.Add(ref template, ref missing, ref missing);

            //Troca o conteúdo de algumas tags
            Word.Word.Range oRng = oDoc.Range(ref missing, ref missing);

            object FindText = "«ETIQUETA»";
            object ReplaceWith = "Teste";
            object MatchWholeWorld = true;
            object Forward = false;


            oRng.Find.Execute(ref FindText, ref missing, ref MatchWholeWorld, ref FindText, ref missing, ref missing, ref Forward, 
                ref missing, ref missing, ref ReplaceWith, ref missing, ref missing, ref missing, ref missing, ref missing);

            oApp.Visible = true;

        }
    }
}
