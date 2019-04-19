using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms.DataVisualization.Charting;

namespace FormulariosAtos
{
    public partial class frm_Main : Form
    {
        public string str_extArquivo;
        public string str_RdoNaoSeguraCarga;
        public string str_RdoMantemCarga;
        public string str_RdoDellPsa;
        public string str_ChkNaoApresProblemas;
        public string str_ChkBatNaoLocalizada;
        public string str_ChkBatFimVida;
        public string str_ChkBatCargaInsuf;
        public string str_ChkBatSemProblema;
        public string str_InsertTblDadosLaudoBat;
        public string str_InsertTblLaudoBat;      

        public frm_Main()
        {
            InitializeComponent();
            txt_EtiquetaCompRoll.Focus();
        }
        
        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;


            wordApp.Selection.Find.Execute(ref ToFindText,
               ref matchCase, ref matchWholeWord,
               ref matchWildCards, ref matchSoundLike,
               ref nmatchAllforms, ref forward,
               ref wrap, ref format, ref replaceWithText,
               ref replace, ref matchKashida,
               ref matchDiactitics, ref matchAlefHamza,
               ref matchControl);
        }

        public void CriaFichaInventario(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //Preenchimento Computador
                this.FindAndReplace(wordApp, "@etiquetacomputador", txt_EtiquetaCompRoll.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@serialcomputador", txt_SerialCompRoll.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@fabricantecomputador", txt_FabCompRoll.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@modelocomputador", txt_ModeloCompRoll.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@dock", txt_EtiquetaDockCompRoll.Text.ToUpper().Trim());

                if (rdo_DesktopRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@tipoequip", "[ X ] DESKTOP [  ] NOTEBOOK          DOCKSTATION: [ X ] NÃO [  ] SIM - ETIQUETA: ");
                }
                else if (rdo_NotebookRoll.Checked == true && chk_DockstationRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@tipoequip", "[  ] DESKTOP [ X ] NOTEBOOK          DOCKSTATION: [  ] NÃO [ X ] SIM - ETIQUETA: " + txt_EtiquetaDockCompRoll.Text);
                }
                else if (rdo_NotebookRoll.Checked == true && chk_DockstationRoll.Checked == false)
                {
                    this.FindAndReplace(wordApp, "@tipoequip", "[  ] DESKTOP [ X ] NOTEBOOK          DOCKSTATION: [ X ] NÃO [  ] SIM - ETIQUETA: ");
                }

                //Preenchimento Monitor
                this.FindAndReplace(wordApp, "@etiquetamonitor", txt_EtiquetaMonitorRoll.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@serialmonitor", txt_SerialMonitorRoll.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@marcamonitor", txt_FabMonitorRoll.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@modelomonitor", txt_ModeloMonitorRoll.Text.ToUpper().Trim());

                //Preenchimento Responsável pelo Equip.
                this.FindAndReplace(wordApp, "@usuarioresponsavel", txt_UsuRespRoll.Text.Trim());
                this.FindAndReplace(wordApp, "@superger", txt_SupUsuRespRoll.Text.ToUpper().Trim() + " - " + txt_GerUsuRespRoll.Text.ToUpper().Trim() + " - " + txt_SetorUsuRespRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@pn", txt_PnUsuRespRoll.Text.Trim());
                this.FindAndReplace(wordApp, "@ramalusuresp", txt_RamalUsuRespRoll.Text.Trim());

                //Preenchimento Dados de Terceiro
                if (chk_EquipCompartilhadoRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@equipcompart", "[ X ] SIM [   ] NÃO");
                }
                else
                {
                    this.FindAndReplace(wordApp, "@equipcompart", "[  ] SIM [ X ] NÃO");
                }
                this.FindAndReplace(wordApp, "@tern", txt_NomeTercRoll.Text.ToUpper().Trim()); //nome
                this.FindAndReplace(wordApp, "@terset", txt_SupTercRoll.Text.ToUpper().Trim() + " - " + txt_GerTercRoll.Text.ToUpper().Trim() + " - " + txt_SetorTercRoll.Text.ToUpper().Trim()); //setor
                this.FindAndReplace(wordApp, "@terra", txt_RamalTercRoll.Text.Trim()); //ramal
                this.FindAndReplace(wordApp, "@matt", txt_MatriculaTercRoll.Text.Trim()); //matricula

                //Preenchimento Localização do Computador
                this.FindAndReplace(wordApp, "@empresa", cbo_EmpresaRoll.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@predio", txt_PredioLocalEquip.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@sala", txt_SalaLocalEquip.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@andar", txt_AndarLocalEquip.Text.Trim() + "º");

                //Preenchimento Informações Complementares
                this.FindAndReplace(wordApp, "@datager", date_FillRoll.Text.Trim());
                this.FindAndReplace(wordApp, "@analresp", cbo_AnalRespRoll.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@numchamado", txt_ChamadoRoll.Text.Trim());
                this.FindAndReplace(wordApp, "@motivoex", txt_MotivoExcRoll.Text.Trim());

                if (chk_MaqExcRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@excecao", "[ X ] SIM [   ] NÃO");
                }
                else
                {
                    this.FindAndReplace(wordApp, "@excecao", "[  ] SIM [ X ] NÃO");
                }

                if (chk_ExcRedeRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@rede", "[ X ] Rede");
                }
                else
                {
                    this.FindAndReplace(wordApp, "@rede", "[  ] Rede");
                }

                if (chk_ExcFisicaRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@fisica", "[ X ] Física");
                }
                else
                {
                    this.FindAndReplace(wordApp, "@fisica", "[  ] Física");
                }
            }
            else
            {
                MessageBox.Show("Arquivo não encontrado");
            }

            //Salvar
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            //Fechar
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("Ficha de inventário gerada com sucesso!");
        }

        private void CriaTermoResp(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //Preenchimento 
                if (rdo_NotebookRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@etiquetaequip", txt_EtiquetaCompRoll.Text.ToUpper().Trim());
                }
                else if (rdo_DesktopRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@etiquetaequip", txt_EtiquetaCompRoll.Text.ToUpper().Trim() + " / " + txt_EtiquetaMonitorRoll.Text.ToUpper().Trim());
                }
                this.FindAndReplace(wordApp, "@empresa", cbo_EmpresaRoll.Text.Trim().ToUpper().Trim());
                this.FindAndReplace(wordApp, "@usuarioresponsavel", txt_UsuRespRoll.Text.Trim().ToUpper().Trim());
                this.FindAndReplace(wordApp, "@pn", txt_PnUsuRespRoll.Text.Trim());
            }
            else
            {
                MessageBox.Show("Arquivo não encontrado");
            }

            //Salvar
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                              ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            //Fechar
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("Termo de responsabilidade gerado com sucesso!!!!");
        }

        private void CriaFormularioDevolucao(object filename, object SaveAs, int tipo)
        {
            //Tipo 0 preenchimento Rollout
            //Tipo 1 preenchimento Devolução

            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                if (tipo == 0)
                {
                    //Preenchimento Rollout
                    this.FindAndReplace(wordApp, "@chamado", txt_ChamadoRoll.Text.Trim());
                    this.FindAndReplace(wordApp, "@analresp", cbo_AnalRespRoll.Text.Trim().ToUpper());
                    this.FindAndReplace(wordApp, "@data", date_FillRoll.Text.Trim());
                    this.FindAndReplace(wordApp, "@nome", txt_UsuRespRoll.Text.Trim());
                    this.FindAndReplace(wordApp, "@pn", txt_PnUsuRespRoll.Text.Trim());
                    this.FindAndReplace(wordApp, "@ramal", txt_RamalUsuRespRoll.Text.Trim());
                    this.FindAndReplace(wordApp, "@gerencia", txt_GerUsuRespRoll.Text.Trim());
                    this.FindAndReplace(wordApp, "@prediosala", txt_PredioLocalEquip.Text.Trim() + " / " + txt_SalaLocalEquip.Text.Trim());
                    this.FindAndReplace(wordApp, "@empresa", cbo_EmpresaDev.Text.ToUpper().Trim());
                }
                else
                {
                    //Preenchimento Devolução
                    this.FindAndReplace(wordApp, "@chamado", txt_ChamadoDev.Text.Trim());
                    this.FindAndReplace(wordApp, "@analresp", cbo_AnalRespDev.Text.Trim().ToUpper());
                    this.FindAndReplace(wordApp, "@data", txt_DataDev.Text.Trim());
                    this.FindAndReplace(wordApp, "@nome", txt_UsuRespEquipDev.Text.Trim());
                    this.FindAndReplace(wordApp, "@pn", txt_PnRespEquipDev.Text.Trim());
                    this.FindAndReplace(wordApp, "@ramal", txt_RamalRespEquipDev.Text.Trim());
                    this.FindAndReplace(wordApp, "@gerencia", txt_GerRespEquipDev.Text.Trim());
                    this.FindAndReplace(wordApp, "@prediosala", txt_PredioLocalEquipDev.Text.Trim() + " / " + txt_SalaLocalEquipDev.Text.Trim());
                    this.FindAndReplace(wordApp, "@empresa", cbo_EmpresaDev.Text.ToUpper().Trim());

                    this.FindAndReplace(wordApp, "@etiqueta1", txt_Etiqueta1Dev.Text.ToUpper().Trim());
                    this.FindAndReplace(wordApp, "@etiqueta2", txt_Etiqueta2Dev.Text.ToUpper().Trim());
                    this.FindAndReplace(wordApp, "@etiqueta3", txt_Etiqueta3Dev.Text.ToUpper().Trim());
                    this.FindAndReplace(wordApp, "@etiqueta4", txt_Etiqueta4Dev.Text.ToUpper().Trim());

                    this.FindAndReplace(wordApp, "@serial1", txt_Serial1Dev.Text.ToUpper().Trim());
                    this.FindAndReplace(wordApp, "@serial2", txt_Serial2Dev.Text.ToUpper().Trim());
                    this.FindAndReplace(wordApp, "@serial3", txt_Serial3Dev.Text.ToUpper().Trim());
                    this.FindAndReplace(wordApp, "@serial4", txt_Serial4Dev.Text.ToUpper().Trim());

                    if (rdo_DesktopDev.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "@desktop", " ( X )Desktop (kit)");
                        this.FindAndReplace(wordApp, "@notebook", "(  )Notebook (kit)");
                        this.FindAndReplace(wordApp, "@monitor", "(  )Somente Monitor");
                        this.FindAndReplace(wordApp, "@periferico", "(  )Periférico Avulso");
                    }
                    else if (rdo_NoteDev.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "@desktop", " (  )Desktop (kit)");
                        this.FindAndReplace(wordApp, "@notebook", "( X )Notebook (kit)");
                        this.FindAndReplace(wordApp, "@monitor", "(  )Somente Monitor");
                        this.FindAndReplace(wordApp, "@periferico", "(  )Periférico Avulso");
                    }
                    else if (rdo_MonitorDev.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "@desktop", " (  )Desktop (kit)");
                        this.FindAndReplace(wordApp, "@notebook", "(  )Notebook (kit)");
                        this.FindAndReplace(wordApp, "@monitor", "( X )Somente Monitor");
                        this.FindAndReplace(wordApp, "@periferico", "(  )Periférico Avulso");
                    }
                    else if (rdo_PerifericoDev.Checked == true)
                    {
                        this.FindAndReplace(wordApp, "@desktop", " (  )Desktop (kit)");
                        this.FindAndReplace(wordApp, "@notebook", "(  )Notebook (kit)");
                        this.FindAndReplace(wordApp, "@monitor", "(  )Somente Monitor");
                        this.FindAndReplace(wordApp, "@periferico", "( X )Periférico Avulso");
                    }
                }
            }
            else
            {
                MessageBox.Show("Arquivo não encontrado");
            }

            //Salvar
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            //Fechar
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("Termo de devolução gerado com sucesso!!!!");
        }

        private void CriaReqPecas(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //Preenchimento 
                this.FindAndReplace(wordApp, "@etiqueta", txt_EtiquetaReqPeca.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@modelo", txt_ModeloReqPeca.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@numserie", txt_SerialReqPeca.Text.ToUpper().Trim());

                this.FindAndReplace(wordApp, "@usuarioresp", txt_UsuReqPeca.Text.Trim().ToUpper().Trim());
                this.FindAndReplace(wordApp, "@pnusuario", txt_PnUsuReqPeca.Text.Trim());
                this.FindAndReplace(wordApp, "@ramal", txt_RamalUsuReqPeca.Text.Trim());

                this.FindAndReplace(wordApp, "@gerente", txt_GerenteReqPeca.Text.Trim().ToUpper().Trim());
                this.FindAndReplace(wordApp, "@pngerente", txt_PnGerenteReqPeca.Text.Trim());

                this.FindAndReplace(wordApp, "@setor", txt_SetorReqPeca.Text.Trim().ToUpper().Trim());
                this.FindAndReplace(wordApp, "@localizacao", cbo_EmpresaRepPeca.Text.Trim());

                this.FindAndReplace(wordApp, "@data", txt_DataReqPeca.Text.Trim());
                this.FindAndReplace(wordApp, "@numchamado", txt_ChamadoReqPeca.Text.Trim());

                this.FindAndReplace(wordApp, "@motivoreparo", txt_CausaReqPeca.Text.Trim());
                this.FindAndReplace(wordApp, "@valorreparo", txt_ValorReqPeca.Text.Trim());

                this.FindAndReplace(wordApp, "@garantiamaq", txt_GarantiaMaqReqPeca.Text.Trim());
                this.FindAndReplace(wordApp, "@garantiaperi", txt_GarantiaPeriReqPeca.Text.Trim());
                this.FindAndReplace(wordApp, "@componente", cbo_ValorReqPeca.Text.ToString().Trim());

                this.FindAndReplace(wordApp, "@gerpn", txt_GerenteReqPeca.Text.Trim().ToUpper().Trim() + " - " + txt_PnGerenteReqPeca.Text.Trim());
                this.FindAndReplace(wordApp, "@usuariopn", txt_UsuReqPeca.Text.Trim().ToUpper().Trim() + " - " + txt_PnUsuReqPeca.Text.Trim());
            }
            else
            {
                MessageBox.Show("Arquivo não encontrado");
            }

            //Salvar
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing);
            //Fechar
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("Formulário de reposição de peças gerado com sucesso!!!!");
        }

        public void CriaLaudoBat(object filename, object SaveAs, string ArqExt)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;
                        
            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing,
                                                   ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //Preenchimento 
                this.FindAndReplace(wordApp, "@usuarioresp", cbo_AnalRespLaudoBat.Text.Trim());
                this.FindAndReplace(wordApp, "@etiqueta", txt_EtiquetaLaudoBat.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@modelo", txt_ModeloLaudoBat.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@numserie", txt_SerialMaqLaudoBat.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@serialbat", txt_SerialBatLaudoBat.Text.ToUpper().Trim());
                this.FindAndReplace(wordApp, "@data", txt_DataLaudoBat.Text.Trim());
                this.FindAndReplace(wordApp, "@numchamado", txt_NumChamadoLaudoBat.Text.Trim());
                this.FindAndReplace(wordApp, "@usuarioresp", txt_UsuRespLaudoBat.Text.Trim().ToUpper().Trim());                
                this.FindAndReplace(wordApp, "@ramal", txt_RamalUsuLaudoBat.Text.Trim());
                this.FindAndReplace(wordApp, "@garantiaequip", txt_GarantiaEquipLaudoBat.Text.Trim());
                this.FindAndReplace(wordApp, "@garantiabat", txt_GarantiaBatLaudoBat.Text.Trim());


                if (str_extArquivo == ".TXT")
                {
                   //this.FindAndReplace(wordApp, "@evidencia", grafico_laudobat.SaveImage("C:\\Documentos ATOS\\Laudobat.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg));
                }
                else
                {
                    //this.FindAndReplace(wordApp, "@evidencia", pic_LaudoBat.);
                }
                


                this.FindAndReplace(wordApp, "@diaglaudo", txt_DiagLaudoBat.Text.Trim().ToUpper().Trim());
                this.FindAndReplace(wordApp, "@solulaudo", txt_SolucaoLaudoBat.Text.Trim().ToUpper().Trim());

                this.FindAndReplace(wordApp, "@dataExtenso", txt_SolucaoLaudoBat.Text.Trim().ToUpper().Trim());
            }        
            else
            {
                MessageBox.Show("Arquivo não encontrado");
            }

            //Salvar
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing);
            //Fechar
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("Formulário de reposição de peças gerado com sucesso!!!!");
        }

        private void ImportaDadosGrafico(string caminho, string tipoarquivo)
        {            
            DataTable dt = new DataTable();
            var lines = File.ReadAllLines(caminho);
            string[] columns = null;
            Regex regex = new Regex(@"^[1-9][0-9]{0,3}$");

            Classes.DataBaseClass ClasseDB = new Classes.DataBaseClass();
            ClasseDB.ConectaAccess();
            ClasseDB.ExecutaQuery("Delete from tbl_laudobat");           
           

            // assuming the first row contains the columns information
            if (lines.Count() > 0)
            {
                columns = lines[3].Split(new char[] { ',' });

               foreach (var column in columns)
                    if (dt.Columns.Count < 4)
                        dt.Columns.Add(column.Trim().ToString().Replace("%", ""));                               
            }

            columns = lines[0].Split(new char[] { ',' });

            // Lendo os dados
            for (int i = 4; i < lines.Count(); i++)
            {
                DataRow dr = dt.NewRow();
                string[] values = lines[i].Split(new char[] { ',' });

                for (int j = 0; j < values.Count() && j < columns.Count(); j++)
                   if (j == 3)
                    {
                        dr[j] = values[j].Replace("%", "");                        
                    }
                    else
                    {
                        dr[j] = values[j];                       
                        //TimeSpan TempoBat = 
                        //txt_DuracLaudoBat.Text = sum.TotalMinutes.ToString();
                    }
                dt.Rows.Add(dr);

                str_InsertTblLaudoBat = "Insert Into tbl_laudobat (datateste_laudobat, horateste_laudobat, status_laudobat, percentual_laudobat) values ('" + values[0].ToString() + "', '" + values[1].ToString() + "', '" + values[2].ToString() + "'," + values[3].Replace("%", "") + ")";

                ClasseDB.ExecutaQuery(str_InsertTblLaudoBat);
            }           
                       
            //Exibe no gráfico

            this.grafico_laudobat.DataSource = dt;
            this.grafico_laudobat.Series["Series1"].XValueMember = "Time";
            this.grafico_laudobat.Series["Series1"].YValueMembers = "Charge";
            this.grafico_laudobat.ChartAreas["ChartArea1"].AxisX.MajorGrid.Enabled = true;
            this.grafico_laudobat.ChartAreas["ChartArea1"].AxisX.Interval = 4;
            this.grafico_laudobat.ChartAreas["ChartArea1"].AxisY.Interval = 20;
            //this.grafico_laudobat.Series["Series1"].IsValueShownAsLabel = true;
            this.grafico_laudobat.DataBind();
            this.grafico_laudobat.Show();
        }

        private void ImportaDadosBateria(string caminho , string tipoarquivo)
    {           
            DataTable dt = new DataTable();
            var lines = File.ReadAllLines(caminho);
            string[] columns = null;

            // Considera a primeira linha como cabeçalho
            if (lines.Count() > 0)
            {
                columns = lines[0].Split(new char[] { ',' });

                foreach (var column in columns)
                    dt.Columns.Add(column);
            }

            //Lê os dados da bateria
            for (int i = 0; i < 3; i++)
            {
                DataRow dr = dt.NewRow();
                string[] values = lines[i].Split(new char[] { ',' });

                for (int j = 0; j < values.Count() && j < columns.Count(); j++)
                    dr[j] = values[j];

                dt.Rows.Add(dr);
            }
            ////Atualiza os dados da bateria
            txt_SerialBatLaudoBat.Text = dt.Rows[1]["Unique ID"].ToString().Trim();
            txt_FabBatLaudoBat.Text = dt.Rows[1][" Manufacturer"].ToString().Trim();
            txt_QuimicaLaudoBat.Text = dt.Rows[1][" Chemistry"].ToString().Trim();
            txt_VoltsLaudoBat.Text = dt.Rows[1][" Voltage (Volts)"].ToString().Trim();            
        }       

        private void frm_Main_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'atosDataSet.tbl_colaboradores' table. You can move, or remove it, as needed.
            this.tbl_colaboradoresTableAdapter.Fill(this.atosDataSet.tbl_colaboradores);
            // TODO: esta linha de código carrega dados na tabela 'atosDataSet.tbl_empresas'. Você pode movê-la ou removê-la conforme necessário.
            this.tbl_empresasTableAdapter.Fill(this.atosDataSet.tbl_empresas);
            // TODO: This line of code loads data into the 'atosDataSet.tbl_reparos' table. You can move, or remove it, as needed.
            this.tbl_reparosTableAdapter.Fill(this.atosDataSet.tbl_reparos);
            txt_ValorReqPeca.Text = "R$" + cbo_ValorReqPeca.SelectedValue.ToString() + ",00";           
        }

        private void chk_DockstationRoll_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_DockstationRoll.Checked == true)
            {
                lbl_EtiquetaDockCompRoll.Enabled = true;
                txt_EtiquetaDockCompRoll.Enabled = true;
            }
            else
            {
                lbl_EtiquetaDockCompRoll.Enabled = false;
                txt_EtiquetaDockCompRoll.Enabled = false;
            }
        }

        private void chk_MaqExcRoll_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_MaqExcRoll.Checked == true)
            {
                txt_MotivoExcRoll.Enabled = true;
                lbl_MotivoExcRoll.Enabled = true;
                chk_ExcFisicaRoll.Enabled = true;
                chk_ExcRedeRoll.Enabled = true;
            }
            else
            {
                txt_MotivoExcRoll.Enabled = false;
                lbl_MotivoExcRoll.Enabled = false;
                chk_ExcFisicaRoll.Enabled = false;
                chk_ExcRedeRoll.Enabled = false;
            }
        }
        
        private void btn_GerarRollout_Click(object sender, EventArgs e)
        {
            if (rdo_Rollout.Checked == true)
            {
                CriaFichaInventario("C:\\Documentos ATOS\\Templates\\FICHA INVENTÁRIO.DOCX", "C:\\Documentos ATOS\\CreatedReports\\FICHA INVENTÁRIO " + txt_ChamadoRoll.Text + ".DOCX");
                CriaTermoResp("C:\\Documentos ATOS\\Templates\\TERMO DE RESPONSABILIDADE.DOC", "C:\\Documentos ATOS\\CreatedReports\\TERMO DE RESPONSABILIDADE " + txt_ChamadoRoll.Text + ".DOC");
                CriaFormularioDevolucao("C:\\Documentos ATOS\\Templates\\TERMO DEVOLUÇÃO DE EQUIPAMENTO.DOCX", "C:\\Documentos ATOS\\CreatedReports\\TERMO DEVOLUÇÃO DE EQUIPAMENTO " + txt_ChamadoRoll.Text + ".DOCX", 0);
            }
            else
            {
                CriaFichaInventario("C:\\Documentos ATOS\\Templates\\FICHA INVENTÁRIO.DOCX", "C:\\Documentos ATOS\\CreatedReports\\FICHA INVENTÁRIO " + txt_ChamadoRoll.Text + ".DOCX");
                CriaTermoResp("C:\\Documentos ATOS\\Templates\\TERMO DE RESPONSABILIDADE.DOC", "C:\\Documentos ATOS\\CreatedReports\\TERMO DE RESPONSABILIDADE " + txt_ChamadoRoll.Text + ".DOC");                
            }
        }

        private void btn_GerarDevolucao_Click(object sender, EventArgs e)
        {
            CriaFormularioDevolucao("C:\\Documentos ATOS\\Templates\\TERMO DEVOLUÇÃO DE EQUIPAMENTO-2.DOCX", "C:\\Documentos ATOS\\CreatedReports\\TERMO DEVOLUÇÃO DE EQUIPAMENTO " + txt_ChamadoDev.Text + ".DOCX", 1);
        }

        private void btn_GerarReqPeca_Click(object sender, EventArgs e)
        {
            CriaReqPecas("C:\\Documentos ATOS\\Templates\\FORMULARIO REPOSIÇÃO DE PEÇAS.DOCX", "C:\\Documentos ATOS\\CreatedReports\\FORMULARIO REPOSIÇÃO DE PEÇAS " + txt_ChamadoReqPeca.Text + ".DOCX");
        }

        private void btn_GerarLaudoBat_Click(object sender, EventArgs e)
        {
            if (txt_NumChamadoLaudoBat.Text == "")
            {
                MessageBox.Show("Favor preencher o número do chamado");
            }
            else
            {
                if (chk_SemCargaLaudoBat.Checked == true)
                {
                    str_RdoNaoSeguraCarga = "[ x ]";
                }
                else
                {
                    str_RdoNaoSeguraCarga = "[    ]";
                }

                if (chk_TestePsaLaudoBat.Checked == true)
                {
                    str_RdoDellPsa = "[ x ]";
                }
                else
                {
                    str_RdoDellPsa = "[    ]";
                }

                if (chk_MantemCargaLaudoBat.Checked == true)
                {
                    str_RdoMantemCarga = "[ x ]";
                }
                else
                {
                    str_RdoMantemCarga = "[    ]";
                }

                if (chk_BateriaCargaInsufLaudoBat.Checked == true)
                {
                    str_ChkBatCargaInsuf = "[ x ]";
                }
                else
                {
                    str_ChkBatCargaInsuf = "[    ]";
                }

                if (chk_BatNaoLocalizadaLaudoBat.Checked == true)
                {
                    str_ChkBatNaoLocalizada = "[ x ]";
                }
                else
                {
                    str_ChkBatNaoLocalizada = "[    ]";
                }

                if (chk_BatCondenadaLaudoBat.Checked == true)
                {
                    str_ChkBatFimVida = "[ x ]";
                }
                else
                {
                    str_ChkBatFimVida = "[    ]";
                }

                if (chk_BatSemProblemaLaudoBat.Checked == true)
                {
                    str_ChkBatSemProblema = "[ x ]";
                }
                else
                {
                    str_ChkBatSemProblema = "[    ]";
                }

                //DateTime DataTeste = DateTime.ParseExact(txt_DataLaudoBat.Text,"dd/MM/yyyy", null);

                string str_msgTesteImpossivel = "Não foi inserido o gráfico com a duração da bateria, pois a bateria do equipamento não está retendo carga suficiente para a realização do teste. Assim que a fonte é desconectada, o notebook desliga em seguida.";

                if (rdo_AnexaArqLaudoBat.Checked == true)
                {
                    str_InsertTblDadosLaudoBat = "Insert Into tbl_dadoslaudobat(numchamado, usuarioresp, localizacao, etiqueta, numserie, serialbat, ramal, datateste, modeloequip, garantiaequip, garantiabat, naoseguracarga, mantemcarga, tempocarga, dellpsa, bateriaok, baterianloc, bateriafimvida, baterianmantem, diagnostico, solucao, analistaresp, cidade, laudo) values('" + txt_NumChamadoLaudoBat.Text + "', '" + txt_UsuRespLaudoBat.Text + "', '" + cbo_LocalLaudoBat.Text + "', '" + txt_EtiquetaLaudoBat.Text + "', '" + txt_SerialMaqLaudoBat.Text + "', '" + txt_SerialBatLaudoBat.Text + "', '" + txt_RamalUsuLaudoBat.Text + "', '" + Convert.ToDateTime(txt_DataLaudoBat.Text) + "', '" + txt_ModeloLaudoBat.Text + "', '" + Convert.ToDateTime(txt_GarantiaEquipLaudoBat.Text) + "', '" + Convert.ToDateTime(txt_GarantiaBatLaudoBat.Text) + "', '" + str_RdoNaoSeguraCarga + "', '" + str_RdoMantemCarga + "', '" + txt_DuracaoBatLaudoBat.Text + "', '" + str_RdoDellPsa + "', '" + str_ChkBatSemProblema + "', '" + str_ChkBatNaoLocalizada + "', '" + str_ChkBatFimVida + "', '" + str_ChkBatCargaInsuf + "', '" + txt_DiagLaudoBat.Text + "', '" + txt_SolucaoLaudoBat.Text + "', '" + cbo_AnalRespLaudoBat.Text + "', '" + cbo_LocalLaudoBat.SelectedValue.ToString() + "', '')";
                }
                else
                {
                    str_InsertTblDadosLaudoBat = "Insert Into tbl_dadoslaudobat(numchamado, usuarioresp, localizacao, etiqueta, numserie, serialbat, ramal, datateste, modeloequip, garantiaequip, garantiabat, naoseguracarga, mantemcarga, tempocarga, dellpsa, bateriaok, baterianloc, bateriafimvida, baterianmantem, diagnostico, solucao, analistaresp, cidade, laudo) values('" + txt_NumChamadoLaudoBat.Text + "', '" + txt_UsuRespLaudoBat.Text + "', '" + cbo_LocalLaudoBat.Text + "', '" + txt_EtiquetaLaudoBat.Text + "', '" + txt_SerialMaqLaudoBat.Text + "', '" + txt_SerialBatLaudoBat.Text + "', '" + txt_RamalUsuLaudoBat.Text + "', '" + Convert.ToDateTime(txt_DataLaudoBat.Text) + "', '" + txt_ModeloLaudoBat.Text + "', '" + Convert.ToDateTime(txt_GarantiaEquipLaudoBat.Text) + "', '" + Convert.ToDateTime(txt_GarantiaBatLaudoBat.Text) + "', '" + str_RdoNaoSeguraCarga + "', '" + str_RdoMantemCarga + "', '" + txt_DuracaoBatLaudoBat.Text + "', '" + str_RdoDellPsa + "', '" + str_ChkBatSemProblema + "', '" + str_ChkBatNaoLocalizada + "', '" + str_ChkBatFimVida + "', '" + str_ChkBatCargaInsuf + "', '" + txt_DiagLaudoBat.Text + "', '" + txt_SolucaoLaudoBat.Text + "', '" + cbo_AnalRespLaudoBat.Text + "', '" + cbo_LocalLaudoBat.SelectedValue.ToString() + "', '" + str_msgTesteImpossivel + "')";
                }                

                Classes.DataBaseClass ClasseDB = new Classes.DataBaseClass();
                ClasseDB.ConectaAccess();
                ClasseDB.ExecutaQuery("Delete from tbl_dadoslaudobat");
                ClasseDB.ExecutaQuery(str_InsertTblDadosLaudoBat);
                ClasseDB.FechaConexao();
                
                frm_CrystalReport FrmRelatorio = new frm_CrystalReport();

                if (rdo_AnexaArqLaudoBat.Checked == true)
                {
                    FrmRelatorio.GeraRelatorioBateria(1);
                    FrmRelatorio.Show();
                }
                else
                {
                    FrmRelatorio.GeraRelatorioBateria(0);
                    FrmRelatorio.Show();
                }
            }            
        }

        private void cbo_ValorReqPeca_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbo_ValorReqPeca.SelectedIndex != -1)
            { 
                txt_ValorReqPeca.Text = "R$" + cbo_ValorReqPeca.SelectedValue.ToString() + ",00";
            }
        }

        private void rdo_NotebookRoll_CheckedChanged(object sender, EventArgs e)
        {
            chk_DockstationRoll.Enabled =       true;
            chk_DockstationRoll.Checked =       false;
        }

        private void rdo_DesktopRoll_CheckedChanged(object sender, EventArgs e)
        {
            chk_DockstationRoll.Enabled =       false;
            txt_EtiquetaDockCompRoll.Enabled =  false;
            lbl_EtiquetaDockCompRoll.Enabled =  false;
        }

        private void btn_LimparRoll_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }

        private void btn_LimparRep_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }

        private void btn_LimparDev_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }

        public void btn_OpenFileLaudoBat_Click(object sender, EventArgs e)
        {
            // open file dialog
            OpenFileDialog open = new OpenFileDialog();

            // Filtro de imagens
            open.Filter = "Arquivos de Texto(*.txt;)|*.txt"; 
                //+ "|Arquivos de Imagens(*.jpg; *.jpeg; *.gif; *.bmp)|*.jpg; *.jpeg; *.gif; *.bmp";


            if (open.ShowDialog() == DialogResult.OK)
            {                
                txt_FileLocationLaudoBat.Text = open.FileName;
                str_extArquivo = Path.GetExtension(txt_FileLocationLaudoBat.Text);
                grafico_laudobat.Visible = true;
                
                ImportaDadosGrafico(txt_FileLocationLaudoBat.Text, str_extArquivo);
                ImportaDadosBateria(txt_FileLocationLaudoBat.Text, str_extArquivo);
            }
        }        

        private void btn_apagaLaudoBat_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }

        private void chk_TestePsaLaudoBat_CheckedChanged(object sender, EventArgs e)
        {
            if(chk_TestePsaLaudoBat.Checked == true)
            {
                chk_BatSemProblemaLaudoBat.Enabled =    true;
                chk_BatNaoLocalizadaLaudoBat.Enabled =  true;
                chk_BatCondenadaLaudoBat.Enabled =      true;
                chk_BateriaCargaInsufLaudoBat.Enabled = true;
            }
            else
            {
                chk_BatSemProblemaLaudoBat.Enabled =    false;
                chk_BatNaoLocalizadaLaudoBat.Enabled =  false;
                chk_BatCondenadaLaudoBat.Enabled =      false;
                chk_BateriaCargaInsufLaudoBat.Enabled = false;
            }
        }

        private void chk_SemCargaLaudoBat_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_SemCargaLaudoBat.Checked == true)
            {
                chk_MantemCargaLaudoBat.Enabled =       false;
                chk_TestePsaLaudoBat.Enabled =          false;
                chk_BatSemProblemaLaudoBat.Enabled =    false;
                chk_BatNaoLocalizadaLaudoBat.Enabled =  false;
                chk_BatCondenadaLaudoBat.Enabled =      false;
                chk_BateriaCargaInsufLaudoBat.Enabled = false;
                txt_DuracaoBatLaudoBat.Enabled =        false;
            }
            else
            {
                chk_MantemCargaLaudoBat.Enabled =       true;
                chk_TestePsaLaudoBat.Enabled =          true;
                chk_BatSemProblemaLaudoBat.Enabled =    true;
                chk_BatNaoLocalizadaLaudoBat.Enabled =  true;
                chk_BatCondenadaLaudoBat.Enabled =      true;
                chk_BateriaCargaInsufLaudoBat.Enabled = true;
                txt_DuracaoBatLaudoBat.Enabled =        true;
            }
        }

        private void rdo_FalhaTesteLaudoBat_CheckedChanged(object sender, EventArgs e)
        {
            if(rdo_FalhaTesteLaudoBat.Checked == true)
            {                
                btn_OpenFileLaudoBat.Enabled =  false;
                grafico_laudobat.Visible =      false;
            }
            else
            {
                rdo_FalhaTesteLaudoBat.Enabled = true;
                btn_OpenFileLaudoBat.Enabled =   true;                
            }
        }
    }
}
