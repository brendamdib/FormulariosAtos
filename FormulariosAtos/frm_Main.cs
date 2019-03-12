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

namespace FormulariosAtos
{
    public partial class frm_Main : Form
    {
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

        private void CriaFichaInventario(object filename, object SaveAs)
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
                this.FindAndReplace(wordApp, "@etiquetacomputador", txt_EtiquetaCompRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@serialcomputador", txt_SerialCompRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@fabricantecomputador", txt_FabCompRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@modelocomputador", txt_ModeloCompRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@dock", txt_EtiquetaDockCompRoll.Text.ToUpper());

                if (rdo_DesktopDev.Checked)
                {                    
                    this.FindAndReplace(wordApp, "@tipoequip", "TIPO DE EQUIPAMENTO: [ X ]DESKTOP [  ]NOTEBOOK      DOCKSTATION: [  ]NÃO [  ]SIM - ETIQUETA: ");
                }
                else if (rdo_NotebookRoll.Checked == true || chk_DockstationRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@tipoequip", "[  ]DESKTOP [ X ]NOTEBOOK      DOCKSTATION: [  ]NÃO [ X ]SIM - ETIQUETA: " + txt_EtiquetaDockCompRoll.Text);
                }
                else if (rdo_NotebookRoll.Checked == true || chk_DockstationRoll.Checked == false)
                {
                    this.FindAndReplace(wordApp, "@tipoequip", "[  ]DESKTOP [ X ]NOTEBOOK      DOCKSTATION: [  ]NÃO [ X ]SIM - ETIQUETA: ");
                }

                //Preenchimento Monitor
                this.FindAndReplace(wordApp, "@etiquetamonitor", txt_EtiquetaMonitorRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@serialmonitor", txt_SerialMonitorRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@marcamonitor", txt_FabMonitorRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@modelomonitor", txt_ModeloMonitorRoll.Text.ToUpper());

                //Preenchimento Responsável pelo Equip.
                this.FindAndReplace(wordApp, "@usuarioresponsavel", txt_UsuRespRoll.Text.Trim());
                this.FindAndReplace(wordApp, "@superger", txt_SupUsuRespRoll.Text.ToUpper() + " - " + txt_GerUsuRespRoll.Text.ToUpper() + " - " + txt_SetorUsuRespRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@pn", txt_PnUsuRespRoll.Text);
                this.FindAndReplace(wordApp, "@ramal", txt_RamalUsuRespRoll.Text);


                //Preenchimento Dados de Terceiro
                this.FindAndReplace(wordApp, "@nometerceiro", txt_NomeTercRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@supergerterceiro", txt_SupTercRoll.Text.ToUpper() + " - " + txt_GerTercRoll.Text.ToUpper() + " - " + txt_SetorTercRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@ramal", txt_RamalTercRoll.Text);
                this.FindAndReplace(wordApp, "@matricul", txt_MatriculaTercRoll.Text);

                //Preenchimento Localização do Computador
                this.FindAndReplace(wordApp, "@empresa", txt_EmpresaLocalEquip.Text.ToUpper());
                this.FindAndReplace(wordApp, "@predio", txt_PredioLocalEquip.Text.ToUpper());
                this.FindAndReplace(wordApp, "@sala", txt_SalaLocalEquip.Text.ToUpper());
                this.FindAndReplace(wordApp, "@andar", txt_AndarLocalEquip.Text + "º" );
                
                //Preenchimento Informações Complementares
                this.FindAndReplace(wordApp, "@datager", date_FillRoll.Text);
                this.FindAndReplace(wordApp, "@analresp", txt_AnalRespRoll.Text.ToUpper());
                this.FindAndReplace(wordApp, "@numchamado", txt_ChamadoRoll.Text);
                this.FindAndReplace(wordApp, "@motivoex", txt_MotivoExcRoll.Text);


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
                    this.FindAndReplace(wordApp, "@etiquetaequip", txt_EtiquetaCompRoll.Text.ToUpper() + " / " + txt_EtiquetaMonitorRoll.Text.ToUpper());
                }
                else if (rdo_DesktopRoll.Checked == true)
                {
                    this.FindAndReplace(wordApp, "@etiquetacomputador", txt_EtiquetaCompRoll.Text.ToUpper());
                }                
                this.FindAndReplace(wordApp, "@empresa", txt_EmpresaLocalEquip.Text.Trim().ToUpper());
                this.FindAndReplace(wordApp, "@usuarioresponsavel", txt_UsuRespRoll.Text.Trim().ToUpper());
                this.FindAndReplace(wordApp, "@pn", txt_PnUsuRespRoll.Text);
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
                    this.FindAndReplace(wordApp, "@chamado", txt_ChamadoRoll.Text);
                    this.FindAndReplace(wordApp, "@analresp", txt_AnalRespRoll.Text.Trim().ToUpper());
                    this.FindAndReplace(wordApp, "@data", date_FillRoll.Text);
                    this.FindAndReplace(wordApp, "@nome", txt_UsuRespRoll.Text);
                    this.FindAndReplace(wordApp, "@pn", txt_PnUsuRespRoll.Text);
                    this.FindAndReplace(wordApp, "@ramal", txt_RamalUsuRespRoll.Text);
                    this.FindAndReplace(wordApp, "@gerencia", txt_GerUsuRespRoll.Text);
                    this.FindAndReplace(wordApp, "@prediosala", txt_PredioLocalEquip.Text + " / " + txt_SalaLocalEquip.Text);
                    this.FindAndReplace(wordApp, "@empresa", txt_EmpresaLocalEquip.Text.ToUpper());
                }
                else
                {
                    //Preenchimento Devolução
                    this.FindAndReplace(wordApp, "@chamado", txt_ChamadoDev.Text);
                    this.FindAndReplace(wordApp, "@analresp", txt_AnalRespDev.Text.Trim().ToUpper());
                    this.FindAndReplace(wordApp, "@data", txt_DataDev.Text);
                    this.FindAndReplace(wordApp, "@nome", txt_NomeRespEquipDev.Text);
                    this.FindAndReplace(wordApp, "@pn", txt_PnRespEquipDev.Text);
                    this.FindAndReplace(wordApp, "@ramal", txt_RamalRespEquipDev.Text);
                    this.FindAndReplace(wordApp, "@gerencia", txt_GerRespEquipDev.Text);
                    this.FindAndReplace(wordApp, "@prediosala", txt_PredioLocalEquipDev.Text + " / " + txt_SalaLocalEquipDev.Text);
                    this.FindAndReplace(wordApp, "@empresa", txt_EmpresaLocalEquipDev.Text.ToUpper());

                    this.FindAndReplace(wordApp, "@etiqueta1", txt_Etiqueta1Dev.Text.ToUpper());
                    this.FindAndReplace(wordApp, "@etiqueta2", txt_Etiqueta2Dev.Text.ToUpper());
                    this.FindAndReplace(wordApp, "@etiqueta3", txt_Etiqueta3Dev.Text.ToUpper());
                    this.FindAndReplace(wordApp, "@etiqueta4", txt_Etiqueta4Dev.Text.ToUpper());                    

                    this.FindAndReplace(wordApp, "@serial1", txt_Serial1Dev.Text.ToUpper());
                    this.FindAndReplace(wordApp, "@serial2", txt_Serial2Dev.Text.ToUpper());
                    this.FindAndReplace(wordApp, "@serial3", txt_Serial3Dev.Text.ToUpper());
                    this.FindAndReplace(wordApp, "@serial4", txt_Serial4Dev.Text.ToUpper());                    

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
                this.FindAndReplace(wordApp, "@etiqueta", txt_EtiquetaReqPeca.Text.ToUpper());
                this.FindAndReplace(wordApp, "@modelo", txt_ModeloReqPeca.Text.ToUpper());
                this.FindAndReplace(wordApp, "@numserie", txt_SerialReqPeca.Text.ToUpper());

                this.FindAndReplace(wordApp, "@usuarioresp", txt_UsuarioReqPeca.Text.Trim().ToUpper());
                this.FindAndReplace(wordApp, "@pnusuario", txt_PnUsuReqPeca.Text);
                this.FindAndReplace(wordApp, "@ramal", txt_RamalUsuReqPeca.Text);

                this.FindAndReplace(wordApp, "@gerente", txt_GerenteReqPeca.Text.Trim().ToUpper());
                this.FindAndReplace(wordApp, "@pngerente", txt_PnGerenteReqPeca.Text);

                this.FindAndReplace(wordApp, "@setor", txt_SetorReqPeca.Text.Trim().ToUpper());
                this.FindAndReplace(wordApp, "@localizacao", txt_LocalizacaoReqPeca.Text);

                this.FindAndReplace(wordApp, "@data", txt_DataReqPeca.Text);
                this.FindAndReplace(wordApp, "@numchamado", txt_ChamadoReqPeca.Text);

                this.FindAndReplace(wordApp, "@motivoreparo", txt_CausaReqPeca.Text);
                this.FindAndReplace(wordApp, "@valorreparo", txt_ValorReqPeca.Text);

                this.FindAndReplace(wordApp, "@garantiamaq", txt_GarantiaMaqReqPeca.Text);
                this.FindAndReplace(wordApp, "@garantiaperi", txt_GarantiaPeriReqPeca.Text);
                this.FindAndReplace(wordApp, "@componente", cbo_ValorReqPeca.SelectedText.ToString());

                this.FindAndReplace(wordApp, "@gerpn", txt_GerenteReqPeca.Text.Trim().ToUpper() + " - " + txt_PnGerenteReqPeca.Text);
                this.FindAndReplace(wordApp, "@usuariopn", txt_UsuarioReqPeca.Text.Trim().ToUpper() + " - " + txt_PnUsuReqPeca.Text);

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
            MessageBox.Show("Formulário de reposição de peças gerado com sucesso!!!!");
        }

        private void frm_Main_Load(object sender, EventArgs e)
        {
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
                CriaFichaInventario("D:\\Documentos ATOS\\Templates\\FICHA INVENTÁRIO.DOCX", "D:\\Documentos ATOS\\FICHA INVENTÁRIO " + txt_ChamadoRoll.Text + ".DOCX");
                CriaTermoResp("D:\\Documentos ATOS\\Templates\\TERMO DE RESPONSABILIDADE.DOC", "D:\\Documentos ATOS\\TERMO DE RESPONSABILIDADE " + txt_ChamadoRoll.Text + ".DOC");
                CriaFormularioDevolucao("D:\\Documentos ATOS\\Templates\\TERMO DEVOLUÇÃO DE EQUIPAMENTO.DOCX", "D:\\Documentos ATOS\\TERMO DEVOLUÇÃO DE EQUIPAMENTO " + txt_ChamadoRoll.Text + ".DOCX", 0);
            }
            else
            {
                CriaFichaInventario("D:\\Documentos ATOS\\Templates\\FICHA INVENTÁRIO.DOCX", "D:\\Documentos ATOS\\FICHA INVENTÁRIO " + txt_ChamadoRoll.Text + ".DOCX");
                CriaTermoResp("D:\\Documentos ATOS\\Templates\\TERMO DE RESPONSABILIDADE.DOC", "D:\\Documentos ATOS\\TERMO DE RESPONSABILIDADE " + txt_ChamadoRoll.Text + ".DOC");                
            }
        }

        private void btn_GerarDevolucao_Click(object sender, EventArgs e)
        {
            CriaFormularioDevolucao("D:\\Documentos ATOS\\Templates\\TERMO DEVOLUÇÃO DE EQUIPAMENTO-2.DOCX", "D:\\Documentos ATOS\\TERMO DEVOLUÇÃO DE EQUIPAMENTO " + txt_ChamadoDev.Text + ".DOCX", 1);
        }

        private void btn_GerarReqPeca_Click(object sender, EventArgs e)
        {
            CriaReqPecas("D:\\Documentos ATOS\\Templates\\FORMULARIO REPOSIÇÃO DE PEÇAS.DOCX", "D:\\Documentos ATOS\\FORMULARIO REPOSIÇÃO DE PEÇAS " + txt_ChamadoReqPeca.Text + ".DOCX");
        }

        private void cbo_ValorReqPeca_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbo_ValorReqPeca.SelectedIndex != -1)
            { 
                txt_ValorReqPeca.Text = "R$" + cbo_ValorReqPeca.SelectedValue.ToString() + ",00";
            }
        }

        private void btn_LimparRoll_Click(object sender, EventArgs e)
        {

        }

        private void rdo_NotebookRoll_CheckedChanged(object sender, EventArgs e)
        {
            chk_DockstationRoll.Enabled =       true;
            chk_DockstationRoll.Checked =       false;

            //Desabilita os campos de monitor
            lbl_FabMonitorRoll.Enabled =        false;
            lbl_ModeloMonitorRoll.Enabled =     false;
            lbl_SerialMonitorRoll.Enabled =     false;
            lbl_EtiquetaMonitorRoll.Enabled =   false;

            txt_FabMonitorRoll.Enabled =        false;
            txt_ModeloMonitorRoll.Enabled =     false;
            txt_SerialMonitorRoll.Enabled =     false;
            txt_EtiquetaMonitorRoll.Enabled =   false;
        }

        private void rdo_DesktopRoll_CheckedChanged(object sender, EventArgs e)
        {
            chk_DockstationRoll.Enabled =       false;
            txt_EtiquetaDockCompRoll.Enabled =  false;
            lbl_EtiquetaDockCompRoll.Enabled =  false;

            //Habilita os campos de monitor
            lbl_FabMonitorRoll.Enabled =        true;
            lbl_ModeloMonitorRoll.Enabled =     true;
            lbl_SerialMonitorRoll.Enabled =     true;
            lbl_EtiquetaMonitorRoll.Enabled =   true;
                                                
            txt_FabMonitorRoll.Enabled =        true;
            txt_ModeloMonitorRoll.Enabled =     true;
            txt_SerialMonitorRoll.Enabled =     true;
            txt_EtiquetaMonitorRoll.Enabled =   true;
        }
    }
}
