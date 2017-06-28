using Busca_LegalDAO;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Systax_BuscaLegal
{
    public partial class Form1 : Form
    {
        WebBrowser webBrowser1;
        DateTime dataCorrente = new DateTime(2013, 10, 03);
        int idUrl = 0;
        int tentativa = 0;
        int countSefaz = 1;
        int itemVez = 0;

        List<string> linksSefazAc;
        List<string> linksSefazCe;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProcessarBusca();

            //var listDocs = new BuscaLegalDao().ObterDocsCorrecao();

            //List<dynamic> listFinal = new List<dynamic>();
            //List<string> listaTratamento2 = new List<string>();
            //List<string> listaErros = new List<string>();

            //string dataReformulada = string.Empty;

            //foreach (var item in listDocs)
            //{
            //    dynamic itemFinal = new ExpandoObject();

            //    itemFinal.Id = item.Id;

            //    string itemCorrigir = item.Texto;

            //    dataReformulada = string.Empty;

            //    itemCorrigir.Replace(" ", "|").Replace("º", string.Empty).Replace("-", "/").ToList().ForEach(delegate(char itemTratar)
            //    {
            //        #region "ID fonte 1,2,7,8,9,10,11,13,14,15"

            //        try
            //        {
            //            if (char.IsNumber(itemTratar) && string.IsNullOrEmpty(dataReformulada))
            //                dataReformulada = itemTratar.ToString();

            //            else if ((char.IsNumber(itemTratar) || "/.".Contains(itemTratar.ToString())) && !string.IsNullOrEmpty(dataReformulada) && !dataReformulada.Contains("|"))
            //                dataReformulada += itemTratar.ToString();

            //            else if (!dataReformulada.Contains("|") && !string.IsNullOrEmpty(dataReformulada))
            //            {
            //                dataReformulada += "|";

            //                if (dataReformulada.Length <= 7)
            //                    dataReformulada = string.Empty;
            //            }
            //        }
            //        catch (Exception)
            //        {
            //        }

            //        #endregion
            //    });

            //    dataReformulada = dataReformulada.Replace("|", string.Empty).Replace(".", "/");

            //    try
            //    {
            //        itemFinal.DataFinal = TratamentoDateTime(dataReformulada).ToString("yyyy-MM-dd");
            //        listFinal.Add(itemFinal);
            //    }
            //    catch (Exception)
            //    {
            //        listaErros.Add(item.Id + ";" + item.Texto);
            //        listaTratamento2.Add(itemCorrigir);
            //    }
            //}

            //#region "CORREÇÂO NÍVEL 2"

            //List<string> listaMeses = new List<string>() { "janeiro|01", "fevereiro|02", "março|03", "marco|03", "abril|04", "maio|05", "junho|06", "julho|07", "agosto|08", "setembro|09", "outubro|10", "novembro|11", "dezembro|12" };

            //listaTratamento2.ForEach(delegate(string itemCorrNivel2)
            //{
            //    var listItensValidar = Regex.Split(itemCorrNivel2, "de ");

            //    if (!string.IsNullOrEmpty(itemCorrNivel2.Trim()))
            //    {
            //        for (int i = 0; i < listItensValidar.Length; i++)
            //        {
            //            try
            //            {
            //                if (listaMeses.Exists(x => x.Split('|')[0].Contains(listItensValidar[i].Trim())))
            //                {
            //                    dataReformulada = Regex.Replace(listItensValidar[i - 1], "[^0-9]+", string.Empty) + "/" + listaMeses.Find(x => x.Split('|')[0].Contains(listItensValidar[i].Trim())).Split('|')[1] + "/" + Regex.Replace(listItensValidar[i + 1], "[^0-9]+", string.Empty);

            //                    dynamic itemFinal = new ExpandoObject();

            //                    itemFinal.Id = int.Parse(listaErros.Find(x => x.Contains(itemCorrNivel2.Trim())).Split(';')[0]);
            //                    itemFinal.DataFinal = TratamentoDateTime(dataReformulada).ToString("yyyy-MM-dd");

            //                    listaErros.RemoveAll(x => x.Contains(itemCorrNivel2.Trim()));
            //                }
            //            }
            //            catch (Exception)
            //            {
            //            }
            //        }
            //    }
            //});

            //#endregion

            //if (listaErros.Count > 0)
            //{
            //    string novoNcm = "ID;data\n";
            //    listaErros.ForEach(x => novoNcm += x + "\n");
            //    File.WriteAllText(@"C:\Temp\ConversaoDateTime.csv", novoNcm);
            //}

            //new BuscaLegalDao().AtualizarDocsCorrecao(listFinal);
        }

        private DateTime TratamentoDateTime(string dataReformulada)
        {
            try
            {
                string ano = dataReformulada.Split('/')[2];
                string mes = dataReformulada.Split('/')[1].Length < 2 ? "0" + dataReformulada.Split('/')[1] : dataReformulada.Split('/')[1];
                string dia = dataReformulada.Split('/')[0].Length < 2 ? "0" + dataReformulada.Split('/')[0] : dataReformulada.Split('/')[0];

                DateTime dataFinal;

                if (dia.Length < 4)
                {
                    if (ano.Length <= 2 && int.Parse(ano) >= int.Parse(DateTime.Now.ToString("yyyy").Substring(2)))
                        ano = "19" + ano;

                    else if (ano.Length <= 2 && int.Parse(ano) <= int.Parse(DateTime.Now.ToString("yyyy").Substring(2)))
                        ano = "20" + ano;

                    dataFinal = new DateTime(int.Parse(ano), int.Parse(mes), int.Parse(dia));
                }
                else
                {
                    dataFinal = new DateTime(int.Parse(dia), int.Parse(mes), int.Parse(ano));
                }

                return dataFinal;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            Thread.Sleep(2000);

            try
            {
                string conteudo = string.Empty;

                do
                {
                    // Do nothing while we wait for the page to load
                }
                while (webBrowser1.ReadyState == WebBrowserReadyState.Loading);

                conteudo = webBrowser1.DocumentText;

                if (conteudo.Contains("Ementas não encontradas. Refine os critérios de pesquisa."))
                {
                    this.BuscarNovaPage();
                }
                else
                {
                    int indexInicio = conteudo.IndexOf("Total de atos localizados");

                    if (indexInicio < 0)
                        ProcessarBusca();

                    else
                    {
                        /*Captura CSS*/
                        var listaCss = Regex.Split(conteudo, "<link").ToList();

                        listaCss.RemoveAt(0);

                        var listaCssTratada = new List<string>();

                        var urlFull = webBrowser1.Url.ToString().Replace("https", "http");

                        urlFull = urlFull.Substring(0, urlFull.IndexOf(".jsf"));

                        listaCss.ForEach(delegate(string x)
                        {
                            string novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                            novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + webBrowser1.Url.ToString().Substring(webBrowser1.Url.ToString().IndexOf("//") + 2).Substring(0, webBrowser1.Url.ToString().Substring(webBrowser1.Url.ToString().IndexOf("//") + 2).IndexOf("/"))));

                            listaCssTratada.Add(novaCss.Replace("\"", string.Empty));
                        });

                        urlFull = string.Empty;

                        listaCssTratada.ForEach(x => urlFull += x);

                        /*Fim Captura CSS*/

                        conteudo = conteudo.Substring(indexInicio);

                        List<string> listaItens = Regex.Split(conteudo.Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " "), "role=row").ToList();

                        listaItens.RemoveRange(0, 2);

                        dynamic estrutura = new ExpandoObject();
                        dynamic itemUrl = new ExpandoObject();
                        dynamic itemListaEmenta;

                        itemUrl.ListaEmenta = new List<dynamic>();

                        listaItens.ForEach(delegate(string y)
                        {
                            try
                            {
                                bool keyTitulo = false;
                                string txtTitulo = " ";

                                y = "<" + y.Substring(0, y.IndexOf("<BR> </span>") + "<BR> </span>".Length);

                                y.ToList().ForEach(delegate(char x) { keyTitulo = x.ToString().Equals("<") || (keyTitulo && !x.ToString().Equals(">")); txtTitulo += !keyTitulo && !x.ToString().Equals(">") ? x.ToString() : txtTitulo.Substring(txtTitulo.Length - 1).Equals("|") ? string.Empty : "|"; });

                                var novaLista = txtTitulo.Split('|').ToList();

                                novaLista.RemoveAll(x => x.Trim() == "");

                                itemListaEmenta = new ExpandoObject();

                                itemListaEmenta.TituloAto = string.Format("{0} {1} Nº{2}, {3}", novaLista[0], novaLista[1], novaLista[2], novaLista[3]);
                                itemListaEmenta.DataEdicao = novaLista[3];
                                itemListaEmenta.Especie = novaLista[0];
                                itemListaEmenta.NumeroAto = novaLista[2];
                                itemListaEmenta.Publicacao = novaLista[3];

                                itemListaEmenta.Republicacao = string.Empty;
                                itemListaEmenta.Escopo = "FED";
                                itemListaEmenta.IdFila = idUrl;

                                y = y.Replace((y.Substring(y.IndexOf("<td role=gridcell"), y.IndexOf("<span style=text-align: justify") - y.IndexOf("<td role=gridcell"))),
                                        "<div class=\"ui-dialog-content ui-widget-content\" style=\"height: 500px;\"><div class=\"line CLASSE_NAO_DEFINIDA\" style=\"ESTILO_NAO_DEFINIDO\"><h1><span style=\"align:center\">" + itemListaEmenta.TituloAto + "</span></h1><br></div><div class=\"line CLASSE_NAO_DEFINIDA\" style=\"ESTILO_NAO_DEFINIDO\"><br><h2>" + "<br></h2></div></div>");

                                itemListaEmenta.Texto = urlFull + removeTagScript(y);
                                itemListaEmenta.Tipo = 3;
                                itemListaEmenta.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                                string textoConteudo = string.Empty;

                                /*Concatena todos os demais valores superiores ao index 4 devido ser conteudo da ementa*/
                                novaLista.GetRange(4, novaLista.Count() - 4).ForEach(x => textoConteudo += x + Environment.NewLine);

                                itemListaEmenta.Hash = this.GerarHash(removerCaracterEspecial(textoConteudo));
                                itemListaEmenta.Ementa = textoConteudo;

                                itemListaEmenta.Sigla = novaLista[1];
                                itemListaEmenta.DescSigla = string.Empty;
                                itemListaEmenta.HasContent = false;

                                itemListaEmenta.ListaArquivos = new List<ArquivoUpload>();

                                itemUrl.ListaEmenta.Add(itemListaEmenta);
                            }
                            catch (Exception ex)
                            {
                                ex.Source = "App|ADE|Loop";
                                new BuscaLegalDao().InserirLogErro(ex, idUrl.ToString(), y);
                            }
                        });

                        itemUrl.IdUrl = idUrl;
                        estrutura.Lista_Nivel2 = new List<dynamic>() { itemUrl };

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { estrutura });
                        this.BuscarNovaPage();
                    }
                }
            }
            catch (Exception ex)
            {
                ex.Source = "App|ADE";
                new BuscaLegalDao().InserirLogErro(ex, idUrl.ToString(), string.Empty);
                this.BuscarNovaPage();
            }
        }

        private void webBrowser1_DocumentCompleted_SefazRs(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                do
                {
                    // Do nothing while we wait for the page to load
                }
                while (webBrowser1.ReadyState == WebBrowserReadyState.Loading);

                if (countSefaz == 1)
                {
                    webBrowser1.Document.GetElementById("RepeaterAreasGrupos_ctl01_cblAreasGrupos_0").InvokeMember("Click");
                    webBrowser1.Document.GetElementById("BtnBuscar").InvokeMember("Click");
                    countSefaz++;
                }
                else
                {
                    var conteudo = webBrowser1.DocumentText;

                    conteudo = conteudo.Substring(conteudo.IndexOf("<ul"), conteudo.IndexOf("</ul>") - conteudo.IndexOf("<ul")).Replace("\"", string.Empty);

                    var listLinks = Regex.Split(conteudo, "<li").ToList();

                    listLinks.RemoveAt(0);

                    dynamic objUrl = new ExpandoObject();

                    objUrl.Indexacao = "Secretaria do estado da fazenda do RS";
                    objUrl.Url = "http://www.legislacao.sefaz.rs.gov.br/";

                    objUrl.Lista_Nivel2 = new List<dynamic>();

                    listLinks.ForEach(delegate(string x)
                    {
                        try
                        {
                            var nrTxt = x.Substring(x.IndexOf("goDocument(") + "goDocument(".Length).Substring(0, x.Substring(x.IndexOf("goDocument(") + "goDocument(".Length).IndexOf(","));

                            nrTxt = "http://www.legislacao.sefaz.rs.gov.br/Site/Document.aspx?inpKey=" + nrTxt + "&inpCodDispositive=&inpDsKeywords=";

                            dynamic itemListaNvl2 = new ExpandoObject();

                            itemListaNvl2.Url = nrTxt;

                            objUrl.Lista_Nivel2.Add(itemListaNvl2);
                        }
                        catch (Exception ex)
                        {
                        }
                    });

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrl });

                    webBrowser1.Document.GetElementById("LinkNext").InvokeMember("Click");
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void webBrowser1_DocumentCompleted_SefazBa(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                do
                {
                    // Do nothing while we wait for the page to load
                }
                while (webBrowser1.ReadyState == WebBrowserReadyState.Loading);

                Thread.Sleep(2000);

                if (countSefaz == 1)
                {
                    countSefaz++;
                    tentativa++;

                    webBrowser1.Document.GetElementById("QueryBox").InnerText = " ";

                    var comboBox = webBrowser1.Document.GetElementById("ViewList");

                    foreach (HtmlElement itemValue in comboBox.Children)
                    {
                        if (itemValue.OuterHtml.ToLower().Contains("tributarioconstituicaobasppublished"))
                        {
                            itemValue.SetAttribute("selected", "selected");
                            break;
                        }

                        else
                            itemValue.SetAttribute("selected", "");
                    }

                    webBrowser1.Document.GetElementById("SearchButton").InvokeMember("Click");
                }
                else
                {
                    countSefaz++;

                    var conteudo = webBrowser1.DocumentText.ToLower().Replace("\"", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty);

                    List<int> xx = new List<int>() { conteudo.IndexOf("<table tabindex=10 id=resultstable"), conteudo.IndexOf("<table id=resultstable tabindex=") };

                    xx.RemoveAll(x => x < 0);

                    conteudo = conteudo.Substring(xx.FirstOrDefault()).Substring(0, conteudo.Substring(xx.FirstOrDefault()).IndexOf("</table>"));

                    var listLinks = Regex.Split(conteudo, "href=").ToList();

                    listLinks.RemoveAt(0);

                    dynamic objUrl = new ExpandoObject();

                    objUrl.Indexacao = "Secretaria do estado da fazenda da Bahia";
                    objUrl.Url = "http://www.sefaz.ba.gov.br/motordebusca/pesquisa/Default.aspx";

                    objUrl.Lista_Nivel2 = new List<dynamic>();

                    if (listLinks.Count > 0)
                    {
                        listLinks.ForEach(delegate(string x)
                                    {
                                        try
                                        {
                                            var nrTxt = x.Substring(0, x.IndexOf("#"));

                                            dynamic itemListaNvl2 = new ExpandoObject();

                                            itemListaNvl2.Url = nrTxt;

                                            objUrl.Lista_Nivel2.Add(itemListaNvl2);
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    });

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrl });
                    }

                    if (webBrowser1.DocumentText.Contains("NextButton"))
                        webBrowser1.Document.GetElementById("NextButton").InvokeMember("Click");
                    else if (tentativa < 6)
                    {
                        string value = "";

                        switch (tentativa)
                        {
                            case 1: value = "tributarioleisestaduaissppublished";
                                break;
                            case 2: value = "tributariodecretossppublished";
                                break;
                            case 3: value = "tributarioportariassppublished";
                                break;
                            case 4: value = "tributarioinstnormativassppublished";
                                break;
                            case 5: value = "consefacordaospdfssppreview";
                                break;
                            default:
                                break;
                        }

                        tentativa++;

                        webBrowser1.Document.GetElementById("QueryBox").InnerText = " ";

                        var comboBox = webBrowser1.Document.GetElementById("ViewList");

                        foreach (HtmlElement itemValue in comboBox.Children)
                        {
                            if (itemValue.OuterHtml.ToLower().Contains(value))
                            {
                                itemValue.SetAttribute("selected", "selected");
                                break;
                            }

                            else
                                itemValue.SetAttribute("selected", "");
                        }

                        webBrowser1.Document.GetElementById("SearchButton").InvokeMember("Click");
                    }
                }
            }
            catch (Exception ex)
            {
                webBrowser1.Document.GetElementById("NextButton").InvokeMember("Click");
            }
        }

        private void webBrowser1_DocumentCompleted_SefazAc(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                do
                {
                    // Do nothing while we wait for the page to load
                }
                while (webBrowser1.ReadyState == WebBrowserReadyState.Loading);

                string htmlUrl = webBrowser1.DocumentText.Replace("\"", string.Empty).Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");

                List<string> listIndicadores = new List<string>();

                listIndicadores.Add(htmlUrl.ToLower().IndexOf("<p>download de constituição estadual atualizada").ToString());
                listIndicadores.Add(htmlUrl.ToLower().IndexOf("<div class=textoconteudo").ToString());

                listIndicadores.RemoveAll(x => x.Equals("-1"));
                listIndicadores = listIndicadores.OrderBy(x => x).ToList();

                if (listIndicadores.Count == 0)
                    listIndicadores = null;

                string linksPdf = htmlUrl.Substring(int.Parse(listIndicadores[0]));

                var listUrls = Regex.Split(linksPdf, "href=").ToList();

                listUrls.RemoveAt(0);
                //listUrls.RemoveAll(x => !x.Contains(".pdf"));

                dynamic objUrll = new ExpandoObject();

                objUrll.Indexacao = "Secretaria do Estado da Fazendo do Acre";
                objUrll.Url = "";

                objUrll.Lista_Nivel2 = new List<dynamic>();

                //Pegando os itens que j são .PDF
                listUrls.FindAll(x => x.Substring(0, x.Trim().IndexOf(" ")).Contains(".pdf")).ToList().ForEach(
                delegate(string itemX)
                {
                    dynamic objItenUrl = new ExpandoObject();

                    objItenUrl.Url = "";
                    objUrll.Lista_Nivel2.Add(objItenUrl);
                });

                listUrls.FindAll(x => !x.Substring(0, x.Trim().IndexOf(" ")).Contains(".pdf") && "?/".Contains(x.Trim().Substring(0, 1))).ToList().ForEach(
                delegate(string itemX)
                {

                });

                new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
            }
        }

        private void webBrowser1_DocumentCompleted_SefazCe(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                do
                {
                    // Do nothing while we wait for the page to load
                }
                while (webBrowser1.ReadyState == WebBrowserReadyState.Loading);

                string htmlUrl = webBrowser1.DocumentText.Replace("\"", string.Empty).Replace("\t", " ").Replace("\n", " ").Replace("\r", " ").Replace("'", string.Empty);

                if (webBrowser1.DocumentText.ToLower().Contains("end of the table of contents has been reached"))
                {
                    itemVez++;
                    ProcessarBusca();
                }
                else
                {
                    List<string> listIndicadores = new List<string>();

                    listIndicadores.Add(htmlUrl.ToLower().IndexOf("<dl><a name={").ToString() + "|</dl>");

                    listIndicadores.RemoveAll(x => x.Equals("-1"));
                    listIndicadores = listIndicadores.OrderBy(x => x).ToList();

                    if (listIndicadores.Count == 0)
                        listIndicadores = null;

                    string linksPdf = htmlUrl.Substring(int.Parse(listIndicadores[0].Split('|')[0])).Substring(0, htmlUrl.Substring(int.Parse(listIndicadores[0].Split('|')[0])).ToLower().IndexOf(listIndicadores[0].Split('|')[1]));

                    var listUrls = Regex.Split(linksPdf.ToLower(), "href=").ToList();

                    listUrls.RemoveAt(0);

                    dynamic objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Secretaria do Estado da Fazendo do Ceará";
                    objUrll.Url = webBrowser1.Url.OriginalString;

                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    listUrls.ForEach(
                    delegate(string itemX)
                    {
                        dynamic objItenUrl = new ExpandoObject();

                        objItenUrl.Url = webBrowser1.Url.OriginalString.Substring(0, webBrowser1.Url.OriginalString.IndexOf(".br") + 3) + itemX.Substring(0, itemX.IndexOf(" "));
                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    });

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });

                    Thread.Sleep(2000);

                    webBrowser1.Document.GetElementById("NEXT").InvokeMember("Click");
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void webBrowser1_DocumentCompletedSefazRo(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            do
            {
                // Do nothing while we wait for the page to load
            }
            while (webBrowser1.ReadyState == WebBrowserReadyState.Loading);

            string htmlUrl = webBrowser1.DocumentText.Replace("\"", string.Empty).Replace("\t", " ").Replace("\n", " ").Replace("\r", " ").Replace("'", string.Empty);
        }

        private void BuscarNovaPage()
        {
            if (dataCorrente <= DateTime.Now)
            {
                if (dataCorrente == new DateTime(2000, 01, 01))
                {
                    webBrowser1.Document.GetElementById("formPrincipal:dataAtoInicial_input").InnerText = dataCorrente.ToString("dd/MM/yyyy");
                    webBrowser1.Document.GetElementById("formPrincipal:dataAtoFinal_input").InnerText = dataCorrente.AddYears(1).ToString("dd/MM/yyyy");

                    dataCorrente = dataCorrente.AddYears(1);
                }
                else
                {
                    webBrowser1.Document.GetElementById("formPrincipal:dataAtoInicial_input").InnerText = dataCorrente.ToString("dd/MM/yyyy");
                    webBrowser1.Document.GetElementById("formPrincipal:dataAtoFinal_input").InnerText = dataCorrente.ToString("dd/MM/yyyy");
                }

                label1.Text = "Processando dia: " + dataCorrente.ToString("dd/MM/yyyy") + " Em: " + DateTime.Now.ToString("HH:mm");

                dataCorrente = dataCorrente.AddDays(1);
                webBrowser1.Document.GetElementById("formPrincipal:btnPesq").InvokeMember("Click");
            }
        }

        private void ProcessarBusca()
        {
            string modoProcessamento = System.Configuration.ConfigurationSettings.AppSettings["modProc"].ToString();

            string siglaFonteProcessamento = System.Configuration.ConfigurationSettings.AppSettings["siglaFonteProcessamento"].ToString();

            List<dynamic> listaUrl = new List<dynamic>();

            #region "Receita Federal - Lote 1"

            List<string> listaCssTratada;

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                string resposta = getHtmlPaginaByGet("http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao", string.Empty);

                /*Nivel 1 - Obter os links do: Acesso direto aos principais atos*/

                resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                var lista = Regex.Split(resposta/*.ToLower()*/.Replace("\"", ""), "outstanding-link").ToList();

                lista.ForEach(delegate(string x)
                {
                    dynamic o = new ExpandoObject();

                    o.Indexacao = x.Contains("Instruções Normativas")
                                    || x.Contains("Atos Declaratórios Interpretativos")
                                        || x.Contains("Pareceres Normativos")
                                            || x.Contains("Portarias") ? "Receita Federal" : null;

                    if (string.IsNullOrEmpty(o.Indexacao))
                        return;

                    o.Url = x.Substring(x.IndexOf("href=") + "href=".Length, x.IndexOf(">") - (x.IndexOf("href=") + "href=".Length));

                    o.Lista_Nivel2 = new List<dynamic>();

                    listaUrl.Add(o);
                });

                dynamic complementoRFB = new ExpandoObject();

                complementoRFB.Indexacao = "Receita Federal";
                complementoRFB.Url = "http://normas.receita.fazenda.gov.br/sijut2consulta/consulta.action?facetsExistentes=&orgaosSelecionados=&tiposAtosSelecionados=7%3B+9%3B+82%3B+10%3B+11%3B+24%3B+30%3B+35%3B+83%3B+38%3B+78%3B+42%3B+79%3B+49%3B+90%3B+77%3B+76%3B+56%3B+54%3B+61%3B+59%3B+57%3B+81%3B+80%3B+66%3B+67%3B+72%3B+75%3B+73&lblTiposAtosSelecionados=AD%3B+ADE%3B+ADEC%3B+ADN%3B+ADI%3B+CO%3B+Dec.%3B+Desp.%3B+DDC%3B+ED%3B+EI%3B+IN%3B+INC%3B+NE%3B+NEC%3B+Nta%3B+NT%3B+OS%3B+ON%3B+Par.%3B+PN%3B+Port.%3B+PC%3B+PIM%3B+Recom.%3B+Res.%3B+SC%3B+SCI%3B+SD&ordemColuna=&ordemDirecao=&tipoAtoFacet=&siglaOrgaoFacet=&anoAtoFacet=&termoBusca=&numero_ato=&tipoData=2&dt_inicio=&dt_fim=&ano_ato=&x=12&y=10&optOrdem=Publicacao_DESC&p=1";
                complementoRFB.Lista_Nivel2 = new List<dynamic>();

                listaUrl.Add(complementoRFB);

                /*Fim Nivel 1 - bter os links do: Acesso direto aos principais atos*/

                /*Nivel 2 - Obter Links da URL relecionada*/

                int count;

                foreach (dynamic itemUrl in listaUrl)
                {
                    count = 1;

                    while (true)
                    {
                        string urlTratada = itemUrl.Url.Replace("&p=1", string.Empty) + "&p=" + count.ToString();

                        string TipoAto = string.Empty;
                        string NrAto = string.Empty;
                        string OrgaoUnid = string.Empty;
                        string Publicacao = string.Empty;
                        string Ementa = string.Empty;

                        try
                        {
                            string respostaNivel2 = getHtmlPaginaByGet(urlTratada, string.Empty);

                            respostaNivel2 = System.Net.WebUtility.HtmlDecode(respostaNivel2/*.ToLower()*/.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            //var listaNivel2 = Regex.Split(respostaNivel2.ToLower().Replace("\"", ""), "gridlinha").ToList();
                            var listaNivel2 = Regex.Split(respostaNivel2.Replace("\"", ""), "gridLinha").ToList();

                            if (respostaNivel2.Contains("Nenhum registro encontrado!"))
                                break;

                            listaNivel2.ForEach(delegate(string x)
                            {
                                if (!x.ToLower().Contains("<th nowrap=") && x.Length >= 10)
                                {
                                    var listaNivel2_1 = Regex.Split(x/*.ToLower()*/.Replace("\"", ""), "<td width=");

                                    int indexCorte = listaNivel2_1[1].IndexOf("href=") + "href=".Length;
                                    string linha = string.Empty;

                                    string url = listaNivel2_1[1].Substring(indexCorte).Substring(0, listaNivel2_1[1].Substring(indexCorte).IndexOf(">"));

                                    url = (itemUrl.Url.Substring(0, itemUrl.Url.LastIndexOf("/") + 1) + url).Replace("'", string.Empty);

                                    linha = listaNivel2_1[1].Substring(indexCorte);
                                    TipoAto = linha.Substring(linha.IndexOf(">") + 1, linha.IndexOf("<") - (linha.IndexOf(">") + 1));

                                    linha = listaNivel2_1[2].Substring(indexCorte);
                                    NrAto = linha.Substring(linha.IndexOf(">") + 1, linha.IndexOf("<") - (linha.IndexOf(">") + 1));

                                    linha = listaNivel2_1[3].Substring(listaNivel2_1[3].IndexOf("title=") + "title=".Length);
                                    OrgaoUnid = linha.Substring(linha.IndexOf(">") + 1, linha.IndexOf("<") - (linha.IndexOf(">") + 1));

                                    linha = listaNivel2_1[4].Substring(indexCorte);
                                    Publicacao = linha.Substring(linha.IndexOf(">") + 1, linha.IndexOf("<") - (linha.IndexOf(">") + 1));

                                    linha = listaNivel2_1[5].Substring(listaNivel2_1[5].IndexOf("href=") + "href=".Length);
                                    Ementa = linha.Substring(linha.IndexOf(">") + 1, linha.IndexOf("<") - (linha.IndexOf(">") + 1));

                                    dynamic itemLista_Nivel2 = new ExpandoObject();

                                    itemLista_Nivel2.Url = url;
                                    itemLista_Nivel2.TipoAto = TipoAto;
                                    itemLista_Nivel2.NrAto = NrAto;
                                    itemLista_Nivel2.OrgaoUnid = OrgaoUnid;
                                    itemLista_Nivel2.Publicacao = Publicacao;
                                    itemLista_Nivel2.Ementa = Ementa;

                                    itemUrl.Lista_Nivel2.Add(itemLista_Nivel2);
                                }
                            });
                        }
                        catch (Exception ex)
                        {
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Format("{0}|{1}|{2}|{3}|{4}|{5}", "Captura URL DOC", TipoAto, NrAto, OrgaoUnid, Publicacao, Ementa));
                        }

                        count++;

                        Thread.Sleep(int.Parse(System.Configuration.ConfigurationSettings.AppSettings["timeSleep"].ToString()));
                    }
                }

                new BuscaLegalDao().AtualizarFontes(listaUrl);

                /*Fim Nivel 2 - Obter Links da URL relecionada*/
            }

            #endregion

            /*Nivel 3 - Obter a Descricao  e outras informações dos Atos*/

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("rfb"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("rfb");

                string aux = string.Empty;
                int posicao = 0;
                string urlTratada = string.Empty;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();

                            /*Estudar como iremos gerar a fila dos docs a partir daqui*/
                            urlTratada = itemLista_Nivel2.Url;

                            itemListaVez.Url = urlTratada;

                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, string.Empty);

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            /*Captura CSS*/
                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            /*Fim Captura CSS*/

                            dynamic dadosEmenta = new ExpandoObject();

                            /*TituloAto*/
                            posicao = respostaNivel2_.IndexOf("<div class=divTituloAto >") + "<div class=divTituloAto >".Length;

                            aux = respostaNivel2_.Substring(posicao);

                            /*Espécie do Ato*/
                            var aux1 = aux.Substring((aux.IndexOf("<div class=left>") + "<div class=left>".Length), (aux.IndexOf("<span style=cursor: default; title=") - (aux.IndexOf("<div class=left>") + "<div class=left>".Length)));

                            dadosEmenta.Especie = aux1.Trim();

                            /*Descricao Sigla do Ato*/
                            var aux2 = aux.Substring(aux.IndexOf("<span style=cursor: default; title=") + "<span style=cursor: default; title=".Length).Substring(0, aux.Substring(aux.IndexOf("<span style=cursor: default; title=") + "<span style=cursor: default; title=".Length).IndexOf(">"));

                            dadosEmenta.DescSigla = aux2.Trim();

                            /*Sigla do Ato*/
                            var aux3 = aux.Substring(aux.IndexOf("<span style=cursor: default; title=") + "<span style=cursor: default; title=".Length).Substring(aux.Substring(aux.IndexOf("<span style=cursor: default; title=") + "<span style=cursor: default; title=".Length).IndexOf(">") + 1, aux.Substring(aux.IndexOf("<span style=cursor: default; title=") + "<span style=cursor: default; title=".Length).IndexOf("</span>") - (aux.Substring(aux.IndexOf("<span style=cursor: default; title=") + "<span style=cursor: default; title=".Length).IndexOf(">") + 1));

                            dadosEmenta.Sigla = aux3.Trim();

                            /*Nº e Data de Edição*/
                            var aux4 = aux.Substring(aux.IndexOf("<!-- FIM: iterator atoOrgaos -->") + "<!-- FIM: iterator atoOrgaos -->".Length, aux.IndexOf("</div>") - (aux.IndexOf("<!-- FIM: iterator atoOrgaos -->") + "<!-- FIM: iterator atoOrgaos -->".Length));

                            dadosEmenta.NumeroAto = aux4.Split(',')[0].Replace("nº", string.Empty).Trim();
                            dadosEmenta.DataEdicao = aux4.Split(',')[1].Trim().Substring(3);

                            /*Anexos existentes no Ato*/

                            var listaArquivos = new List<string>();

                            dadosEmenta.ListaArquivos = new List<ArquivoUpload>();

                            if (respostaNivel2_.IndexOf("anexoOutros.action?idArquivoBinario") >= 0)
                            {
                                string aux5 = respostaNivel2_.Substring(respostaNivel2_.IndexOf("anexoOutros.action?idArquivoBinario"));
                                int indexNovo = 0;

                                while (aux5.IndexOf("anexoOutros.action?idArquivoBinario=") >= 0)
                                {
                                    indexNovo = aux5.IndexOf("anexoOutros.action?idArquivoBinario=");

                                    aux5 = aux5.Substring(indexNovo + "anexoOutros.action?idArquivoBinario=".Length);

                                    if (!aux5.Substring(0, aux5.IndexOf("</a>")).Split('>')[0].Equals("0"))
                                        listaArquivos.Add("anexoOutros.action?idArquivoBinario=" + aux5.Substring(0, aux5.IndexOf("</a>")).Split('>')[0] + "|" + aux5.Substring(0, aux5.IndexOf("</a>")).Split('>')[1]);
                                }

                                listaArquivos.ForEach(delegate(string x)
                                {
                                    string nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), x.Split('|')[1]);

                                    using (WebClient webClient = new WebClient())
                                    {
                                        webClient.DownloadFile(urlTratada.Substring(0, urlTratada.LastIndexOf("/")) + "/" + x.Split('|')[0], nomeArq);
                                    }

                                    string dadosPdf = LeArquivo(nomeArq);

                                    byte[] arrayFile = File.ReadAllBytes(nomeArq);

                                    dadosEmenta.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = nomeArq.Substring(nomeArq.LastIndexOf(".") + 1), NomeArquivo = x.Split('|')[1] });

                                    File.Delete(nomeArq);
                                });
                            }

                            aux = aux.Substring(0, aux.IndexOf("<!-- INICIO: Menu - Vis"));

                            string txtTitulo = string.Empty;
                            bool keyTitulo = false;

                            /*Isolar essa parte em uma nova função*/
                            aux.ToList().ForEach(delegate(char x) { keyTitulo = x.ToString().Equals("<") || (keyTitulo && !x.ToString().Equals(">")); txtTitulo += !keyTitulo && !x.ToString().Equals(">") ? x.ToString() : string.Empty; });

                            dadosEmenta.TituloAto = Regex.Replace(txtTitulo.Trim(), @"\s+", " ");

                            /*Publicacao*/
                            posicao = respostaNivel2_.IndexOf("class=left cleft") + "class=left cleft".Length;

                            aux = respostaNivel2_.Substring(posicao);
                            aux = aux.Substring(aux.IndexOf(">") + 1);
                            aux = aux.Substring(0, aux.IndexOf("<"));

                            dadosEmenta.Publicacao = aux;
                            dadosEmenta.Republicacao = string.Empty;
                            dadosEmenta.Escopo = "FED";
                            dadosEmenta.IdFila = itemLista_Nivel2.Id;

                            /*Ementa*/
                            posicao = respostaNivel2_.IndexOf("<div id=divTexto>") + "<div id=divTexto>".Length;

                            posicao = (respostaNivel2_.Substring(posicao).Contains("p class=ementa") ? respostaNivel2_.IndexOf("p class=ementa") + "p class=ementa".Length : posicao);

                            aux = respostaNivel2_.Substring(posicao).Replace("<strike>", string.Empty).Replace("</strike>", string.Empty);
                            aux = aux.Substring(aux.IndexOf(">") + 1);
                            aux = aux.Substring(0, (aux.IndexOf("</p>") >= 0 ? aux.IndexOf("</p>") : aux.IndexOf("</span>")));

                            /*Isolar essa parte em uma nova função*/
                            bool key = false;
                            string texto = string.Empty;

                            /*Isolar essa parte em uma nova função*/
                            aux.ToList().ForEach(delegate(char x) { key = x.ToString().Equals("<") || (key && !x.ToString().Equals(">")); texto += !key && !x.ToString().Equals(">") ? x.ToString() : string.Empty; });

                            dadosEmenta.Ementa = texto;

                            /*Criacao Metadado e Hash*/

                            dadosEmenta.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                            //dadosEmenta.Hash = this.GerarHash(dadosEmenta.Metadado);

                            dadosEmenta.HasContent = false;
                            /*Fim - Criacao Metadado e Hash*/

                            /*Texto*/
                            dadosEmenta.Texto = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<div class=divTituloAto >"), respostaNivel2_.LastIndexOf("class=divSegmentos") - respostaNivel2_.IndexOf("<div class=divTituloAto >"));

                            dadosEmenta.Texto = removeTagScript(dadosEmenta.Texto);

                            dadosEmenta.Hash = this.GerarHash(removerCaracterEspecial(obterTextoConteudo(dadosEmenta.Texto + ">", 3)));

                            dadosEmenta.Texto = novaCss + dadosEmenta.Texto;

                            /*Tipo Texto*/
                            dadosEmenta.Tipo = 3;

                            itemListaVez.ListaEmenta.Add(dadosEmenta);

                            /*Obtendo as demais variações do Texto*/

                            /*1 - Original*/
                            respostaNivel2_ = getHtmlPaginaByGet(urlTratada.Replace("anotado", "original"), string.Empty);

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_/*.ToLower()*/.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            /*Validar se o marca-inico e fim são iguais*/
                            respostaNivel2_ = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<div class=divTituloAto >"), respostaNivel2_.LastIndexOf("class=divSegmentos") - respostaNivel2_.IndexOf("<div class=divTituloAto >"));
                            var hashComparacao = this.GerarHash(removerCaracterEspecial(obterTextoConteudo(removeTagScript(respostaNivel2_), 1)));

                            /*Caso conteúdos dos documentos forem iguais não inserimos as variações dos documentos*/
                            if (!hashComparacao.Equals(dadosEmenta.Hash))
                            {
                                dadosEmenta = new ExpandoObject();

                                dadosEmenta.ListaArquivos = new List<ArquivoUpload>();

                                dadosEmenta.Texto = novaCss + respostaNivel2_;
                                dadosEmenta.Hash = hashComparacao;
                                dadosEmenta.HasContent = false;
                                dadosEmenta.Tipo = 1;

                                itemListaVez.ListaEmenta.Add(dadosEmenta);

                                /*2 - Vigente*/
                                respostaNivel2_ = getHtmlPaginaByGet(urlTratada.Replace("anotado", "compilado"), string.Empty);

                                respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_/*.ToLower()*/.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                                /*Validar se o marca-inico e fim são iguais*/
                                respostaNivel2_ = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<div class=divTituloAto >"), respostaNivel2_.LastIndexOf("class=divSegmentos") - respostaNivel2_.IndexOf("<div class=divTituloAto >"));

                                dadosEmenta = new ExpandoObject();

                                dadosEmenta.ListaArquivos = new List<ArquivoUpload>();

                                dadosEmenta.Texto = novaCss + respostaNivel2_;
                                dadosEmenta.Hash = this.GerarHash(removerCaracterEspecial(obterTextoConteudo(respostaNivel2_, 2)));
                                dadosEmenta.HasContent = false;
                                dadosEmenta.Tipo = 2;

                                itemListaVez.ListaEmenta.Add(dadosEmenta);
                            }

                            dynamic listaNivel2Nova = new ExpandoObject();

                            listaNivel2Nova.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { listaNivel2Nova });

                            Thread.Sleep(int.Parse(System.Configuration.ConfigurationSettings.AppSettings["timeSleep"].ToString()));
                        }
                        catch (Exception ex)
                        {
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Format("{0}", "Captura DOC"));
                        }
                    }
                }
            }

            #endregion

            /*Fim Nivel 3 - Obter a Descricao  e outras informações dos Atos*/

            #endregion

            #region "Comissão de Valores Mobiliários - Lote 2"

            HttpWebRequest request = null;

            int countCVM = 1;
            string auxCVM = string.Empty;
            string urlCVM = string.Empty;

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                dynamic itemListaNvl2;

                while (true)
                {
                    urlCVM = string.Empty;

                    try
                    {
                        request = (HttpWebRequest)WebRequest.Create("http://www.cvm.gov.br/legislacao/index.html?numero=&lastNameShow=&lastName=&filtro=todos&dataInicio=&dataFim=&categoria0=%2Flegislacao%2Finst%2F&categoria1=%2Flegislacao%2Fpare%2F&categoria2=%2Flegislacao%2Fdeli%2F&categoria3=%2Flegislacao%2Fconj%2F&categoria4=%2Flegislacao%2Fcirc%2F&buscado=false&contCategoriasCheck=7");

                        var postData = "searchPage=" + countCVM.ToString();
                        postData += "&lastName=";
                        postData += "&numero=";
                        postData += "&filtro=";
                        postData += "&todos";
                        postData += "&dataInicio=";
                        postData += "&dataFim=";
                        postData += "&buscado=false";
                        postData += "&contCategoriasCheck=7";
                        postData += "&itensPagina=10";
                        postData += "&ordenar=recentes";
                        postData += "&dataInicioBound=1978%2F04%2F27";
                        postData += string.Format("&dataFimBound={0}%2F{1}%2F{2}", DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"));
                        postData += "&listaBuscaAside=";
                        postData += "&tipos=";

                        var data = Encoding.ASCII.GetBytes(postData);

                        request.Method = "POST";
                        request.ContentType = "application/x-www-form-urlencoded";
                        request.ContentLength = data.Length;

                        using (var stream = request.GetRequestStream())
                        {
                            stream.Write(data, 0, data.Length);
                        }

                        var response = (HttpWebResponse)request.GetResponse();

                        var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                        if (!responseString.ToLower().Contains("<article>") || countCVM > 1000)
                            break;

                        responseString = responseString.Substring(responseString.IndexOf("<article>"));

                        responseString = System.Net.WebUtility.HtmlDecode(responseString.Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                        //Loop nos itens dentro da paginação via Post.

                        dynamic objUrl = new ExpandoObject();

                        objUrl.Indexacao = "Comissão de valores mobiliários";
                        objUrl.Url = "http://www.cvm.gov.br/legislacao";

                        objUrl.Lista_Nivel2 = new List<dynamic>();

                        foreach (var itemCVM in Regex.Split(responseString.Replace("\"", string.Empty), "<article>"))
                        {
                            try
                            {
                                if (!string.IsNullOrEmpty(itemCVM))
                                {
                                    auxCVM = itemCVM.Substring(itemCVM.IndexOf("</h3>") + "</h3>".Length);
                                    auxCVM = auxCVM.Substring(auxCVM.IndexOf(">") + 1).Substring(0, auxCVM.Substring(auxCVM.IndexOf(">") + 1).IndexOf("<")).Trim();

                                    string nomeInstrucao = auxCVM;

                                    auxCVM = string.Empty;
                                    auxCVM = itemCVM.Substring(itemCVM.IndexOf("href=") + "href=".Length).Substring(0, itemCVM.Substring(itemCVM.IndexOf("href=") + "href=".Length).IndexOf("title")).Trim();

                                    urlCVM = "http://www.cvm.gov.br" + auxCVM;

                                    itemListaNvl2 = new ExpandoObject();

                                    itemListaNvl2.Url = urlCVM;

                                    objUrl.Lista_Nivel2.Add(itemListaNvl2);
                                    //Continuar a implementação do salvar da URL
                                }
                            }
                            catch (Exception ex)
                            {
                                new BuscaLegalDao().InserirLogErro(ex, urlCVM, string.Format("{0}", "1.1 Captura Url - CVM"));
                            }
                        }

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrl });

                        countCVM++;
                    }
                    catch (Exception ex)
                    {
                        new BuscaLegalDao().InserirLogErro(ex, urlCVM, string.Format("{0}", "1 Captura Url - CVM"));
                    }
                }
            }

            #endregion

            #region "Captura DOC's"

            /*Nível Obter a Url -> Obter link do arquivo e o download do arquivo*/
            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("cvm"))
            {
                this.Text = "Fonte: CVM";

                this.label1.Text = "Processando";

                try
                {
                    List<dynamic> listaProcessamento = new BuscaLegalDao().ObterUrlsParaProcessamento("cvm");

                    foreach (dynamic itemUrl in listaProcessamento.FirstOrDefault().Lista_Nivel2)
                    {
                        try
                        {
                            WebRequest req = WebRequest.Create(itemUrl.Url);
                            req.Headers.Add("Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.6,en;q=0.4");
                            WebResponse res = req.GetResponse();
                            Stream dataStream = res.GetResponseStream();
                            StreamReader reader = new StreamReader(dataStream);
                            string resposta = reader.ReadToEnd();

                            //string resposta = getHtmlPaginaByGet(itemUrl.Url, string.Empty);

                            resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ")).Replace("\"", " ");

                            auxCVM = resposta.Substring(resposta.IndexOf("class=download"));

                            urlCVM = "http://www.cvm.gov.br" + auxCVM.Substring(auxCVM.IndexOf("href=") + "href=".Length).Substring(0, auxCVM.Substring(auxCVM.IndexOf("href=") + "href=".Length).IndexOf("title")).Trim();

                            string nomeArq = urlCVM.Substring(urlCVM.LastIndexOf("/") + 1);

                            auxCVM = resposta.Substring(resposta.IndexOf("<article class=contentTextoGeral>")).Substring(0, resposta.Substring(resposta.IndexOf("<article class=contentTextoGeral>")).IndexOf("<div class=contentVejaMais pull-left no-border no-padding>"));

                            string txtTitulo = string.Empty;
                            bool keyTitulo = false;

                            /*Isolar essa parte em uma nova função*/
                            List<string> ListaTratada = new List<string>();

                            Regex.Split(auxCVM.Trim().Replace("<p class=MsoNormal>", "<p>").Replace("<p style=text-align: justify;>", "<p>").Replace("<p style=text-align: left;>", "<p>"), "<p>")
                                 .ToList()
                                 .ForEach(delegate(string x)
                                 {
                                     txtTitulo = string.Empty; x.ToList()
                                                                .ForEach(delegate(char y)
                                                                {
                                                                    keyTitulo = y.ToString().Equals("<") || (keyTitulo && !y.ToString().Equals(">")); txtTitulo += !keyTitulo && !y.ToString().Equals(">") ? y.ToString() : string.Empty;
                                                                });

                                     ListaTratada.Add(txtTitulo.Trim());
                                 });

                            //txtTitulo = Regex.Replace(txtTitulo.Trim(), @"\s+", " ");

                            nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(urlCVM, nomeArq);
                            }

                            string dadosPdf = LeArquivo(nomeArq);

                            //byte[] arrayFile = File.ReadAllBytes(nomeArq);

                            File.Delete(nomeArq);

                            itemUrl.ListaEmenta = new List<dynamic>();

                            dynamic itemListaEmenta = new ExpandoObject();

                            itemListaEmenta.TituloAto = ListaTratada[0].Trim();

                            itemListaEmenta.DataEdicao = ListaTratada[1].Trim().Substring(0, 10);

                            /*Tramento Número*/
                            var indexNumCorte = Regex.Replace(itemListaEmenta.TituloAto, "[^0-9]+", string.Empty).Substring(0, 1);
                            itemListaEmenta.NumeroAto = itemListaEmenta.TituloAto.Substring(itemListaEmenta.TituloAto.IndexOf(indexNumCorte)).Contains("/") ?
                                                            itemListaEmenta.TituloAto.Substring(itemListaEmenta.TituloAto.IndexOf(indexNumCorte)).Substring(0, itemListaEmenta.TituloAto.Substring(itemListaEmenta.TituloAto.IndexOf(indexNumCorte)).IndexOf("/")) :
                                                                itemListaEmenta.TituloAto;

                            itemListaEmenta.NumeroAto = Regex.Replace(itemListaEmenta.NumeroAto, "[^0-9]+", string.Empty);

                            itemListaEmenta.Especie = itemListaEmenta.TituloAto.Split(' ')[0];

                            /*Modelo para a classe DAO*/
                            itemListaEmenta.ListaArquivos = new List<ArquivoUpload>();

                            itemListaEmenta.Publicacao = string.Empty;
                            ListaTratada.ForEach(x => itemListaEmenta.Publicacao = itemListaEmenta.Publicacao.Equals(string.Empty) && x.Contains("(Publicad") ? x.Substring(0, (x.IndexOf(")") + 1)) : itemListaEmenta.Publicacao);

                            if (itemListaEmenta.Publicacao.Equals(string.Empty))
                                ListaTratada.ForEach(x => itemListaEmenta.Publicacao = itemListaEmenta.Publicacao.Equals(string.Empty) && x.Substring(2, 1).Equals("/") ? x.Trim().Substring(0, 10) : itemListaEmenta.Publicacao);

                            itemListaEmenta.Republicacao = string.Empty;
                            itemListaEmenta.Escopo = "FED";
                            itemListaEmenta.IdFila = itemUrl.Id;

                            itemListaEmenta.Ementa = (ListaTratada.Count > 2 ? ListaTratada[2] : Regex.Split(ListaTratada[1], "     ").Last().Trim());
                            itemListaEmenta.Texto = dadosPdf;
                            itemListaEmenta.Tipo = 3;
                            itemListaEmenta.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";
                            itemListaEmenta.Hash = this.GerarHash(removerCaracterEspecial(dadosPdf));

                            itemListaEmenta.Sigla = string.Empty;
                            itemListaEmenta.DescSigla = string.Empty;

                            itemListaEmenta.HasContent = false;
                            //itemListaEmenta.ByteArrPdf = arrayFile;
                            itemListaEmenta.NomeArquivo = urlCVM.Substring(urlCVM.LastIndexOf("/") + 1);
                            itemListaEmenta.ExtensaoArquivo = urlCVM.Substring(urlCVM.LastIndexOf(".") + 1);

                            itemUrl.ListaEmenta.Add(itemListaEmenta);

                            dynamic estrutura = new ExpandoObject();

                            estrutura.Lista_Nivel2 = new List<dynamic>() { itemUrl };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { estrutura });
                            Thread.Sleep(int.Parse(System.Configuration.ConfigurationSettings.AppSettings["timeSleep"].ToString()));
                        }
                        catch (Exception ex)
                        {
                            new BuscaLegalDao().InserirLogErro(ex, urlCVM, string.Format("{0}", "1.1 Captura Doc - CVM"));
                        }
                    }

                    this.label1.Text = "Concluído";
                }
                catch (Exception ex)
                {
                    new BuscaLegalDao().InserirLogErro(ex, urlCVM, string.Format("{0}", "1 Captura Doc - CVM"));
                }
            }

            #endregion

            #endregion

            #region "Decisões Administrativas - Lote 3"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                dynamic objUrl = new ExpandoObject();

                objUrl.Indexacao = "Atos Decisórios - Ementários";
                objUrl.Url = "https://novodecisoes.receita.fazenda.gov.br/consultaweb/index.jsf";

                objUrl.Lista_Nivel2 = new List<dynamic>();

                dynamic itemListaNvl2 = new ExpandoObject();

                itemListaNvl2.Url = objUrl.Url;

                objUrl.Lista_Nivel2.Add(itemListaNvl2);

                new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrl });
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("ade"))
            {
                this.Text = "Fonte: Atos Normativos - Ementários";

                webBrowser1 = new WebBrowser();

                if (idUrl == 0)
                {
                    List<dynamic> listaProcessamento = new BuscaLegalDao().ObterUrlsParaProcessamento("ade");
                    idUrl = listaProcessamento.FirstOrDefault().Lista_Nivel2[0].IdUrl;
                }

                webBrowser1.ScriptErrorsSuppressed = true;

                webBrowser1.Navigate(@"http://novodecisoes.receita.fazenda.gov.br/consultaweb/index.jsf");

                webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted);
            }

            #endregion

            #endregion

            #region "Receita Federal(Legislação) - Lote 3"

            /*Deixar independente o processamento desses documento, remover o vinculo com a RFB*/

            #region "Captura URL's"

            List<dynamic> listaInserir;
            List<string> processarList;

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                listaInserir = new List<dynamic>();

                processarList = new List<string>(){"http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/medidas-provisorias"
                                                  ,"http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/leis-complementares"
                                                  ,"http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/decretos-leis"
                                                  ,"http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/leis"
                                                  ,"http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/link"
                                                  ,"http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/regulamentos"};

                processarList.ForEach(x => listaInserir.Add(obterEstruturaProcessarUrl_RfbLeg(x)));

                new BuscaLegalDao().AtualizarFontes(listaInserir);
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("rfb-leg"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("rfb-leg");

                string urlTratada = string.Empty;

                int saiFora = 0;

                dynamic ementa;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            saiFora++;

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim();

                            itemListaVez.Url = urlTratada;

                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            WebRequest reqNivel2_ = WebRequest.Create(urlTratada);

                            reqNivel2_.Headers.Add("Accept-Encoding", "gzip, deflate, br");
                            reqNivel2_.Headers.Add("Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.6,en;q=0.4");

                            WebResponse resNivel2_ = reqNivel2_.GetResponse();
                            Stream dataStreamNivel2_ = resNivel2_.GetResponseStream();
                            StreamReader readerNivel2_;

                            //if (!urlTratada.ToLower().Contains("/leis/"))
                            readerNivel2_ = new StreamReader(dataStreamNivel2_, System.Text.Encoding.Default);
                            //else
                            //readerNivel2_ = new StreamReader(dataStreamNivel2_);

                            string respostaNivel2_ = readerNivel2_.ReadToEnd();

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            /*Captura CSS*/
                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            string novaCss = string.Empty;

                            if (listaCss.Count > 0)
                            {
                                listaCssTratada = new List<string>();

                                listaCss.ForEach(delegate(string x)
                                {
                                    novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                    listaCssTratada.Add(novaCss);
                                });

                                novaCss = string.Empty;

                                listaCssTratada.ForEach(x => novaCss += x);
                            }

                            /*Fim Captura CSS*/

                            ementa = new ExpandoObject();

                            /**Titulo**/
                            string correcao = respostaNivel2_.ToLower().Replace("https", "http").Replace("'", string.Empty);
                            var titulos = correcao.Substring(correcao.IndexOf("href=http://legislacao.planalto.gov.br/legisla/"));

                            respostaNivel2_ = correcao;

                            int indexCorteTitle = titulos.IndexOf("</strong>") < 0 ? titulos.IndexOf("</small>") : titulos.IndexOf("</strong>");

                            titulos = ("<" + (titulos.Substring(0, indexCorteTitle)));

                            ementa.TituloAto = ObterStringLimpa(titulos).Trim().ToLower();

                            /*Melhorar esse tratamento*/
                            if (urlTratada.ToLower().Contains("/decreto-lei/") || urlTratada.ToLower().Contains("/decreto/") || urlTratada.ToLower().Contains("/leis/") || urlTratada.ToLower().Contains("/lei/"))
                            {
                                /**Data**/
                                ementa.DataEdicao = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" de") + 4, (ementa.TituloAto.Length) - (ementa.TituloAto.IndexOf(" de") + 4));
                                ementa.Publicacao = ementa.DataEdicao;
                            }
                            else
                            {
                                /**Data**/
                                ementa.DataEdicao = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" de") + 3, (ementa.TituloAto.Length) - (ementa.TituloAto.IndexOf(" de") + 3));
                                ementa.Publicacao = ementa.DataEdicao;
                            }

                            try
                            {
                                /**Espécie**/
                                ementa.Especie = ementa.TituloAto.Substring(0, ementa.TituloAto.IndexOf(" n")).Trim();
                                /**Numero**/
                                //ementa.NumeroAto = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).Substring(0, ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).IndexOf(",")).Trim();
                                ementa.NumeroAto = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).Substring(0, ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).IndexOf(",") >= 0 ? ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).IndexOf(",") : ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).IndexOf("de")).Trim();
                            }
                            catch (Exception)
                            {
                                ementa.Especie = ementa.TituloAto.Substring(0, ementa.TituloAto.IndexOf("n")).Trim();

                                //ementa.NumeroAto = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).Substring(0, ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).IndexOf(",")).Trim();
                                ementa.NumeroAto = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).Substring(0, ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).IndexOf(",") >= 0 ? ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).IndexOf(",") : ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).IndexOf("de")).Trim();
                            }

                            ementa.NumeroAto = Regex.Replace(ementa.NumeroAto, "[^0-9]+", string.Empty);
                            ementa.Tipo = 3;

                            /**Metadado**/
                            ementa.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                            /**Outros**/
                            ementa.Sigla = string.Empty;
                            ementa.DescSigla = string.Empty;
                            ementa.HasContent = false;
                            ementa.Republicacao = string.Empty;

                            ementa.Escopo = "FED";
                            ementa.IdFila = itemLista_Nivel2.Id;

                            /**Arquivo**/
                            ementa.ListaArquivos = new List<ArquivoUpload>();

                            if (respostaNivel2_.IndexOf("de Motivos") > 0)
                            {
                                int indexArq = 12;
                                string nomeArquivo = string.Empty;

                                while (!respostaNivel2_.Substring(respostaNivel2_.IndexOf("de Motivos") - indexArq).Substring(0, 1).Equals("="))
                                {
                                    nomeArquivo = respostaNivel2_.Substring(respostaNivel2_.IndexOf("de Motivos") - indexArq).Substring(0, 1) + nomeArquivo;
                                    indexArq++;
                                }

                                if (!nomeArquivo.Substring(nomeArquivo.LastIndexOf(".") + 1).Contains("htm"))
                                {
                                    urlTratada = urlTratada.Substring(0, urlTratada.LastIndexOf("/")).Substring(0, urlTratada.Substring(0, urlTratada.LastIndexOf("/")).LastIndexOf("/"));
                                    nomeArquivo = nomeArquivo.Substring(nomeArquivo.IndexOf("/"));

                                    string nomeArqFull = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArquivo.Substring(nomeArquivo.LastIndexOf("/") + 1));

                                    using (WebClient webClient = new WebClient())
                                    {
                                        webClient.DownloadFile(string.Format("{0}{1}", urlTratada, nomeArquivo), nomeArqFull);
                                    }

                                    byte[] arrayFile = File.ReadAllBytes(nomeArqFull);

                                    ementa.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = nomeArquivo.Substring(nomeArquivo.LastIndexOf(".") + 1), NomeArquivo = nomeArquivo.Substring(nomeArquivo.LastIndexOf("/") + 1) });

                                    File.Delete(nomeArqFull);
                                }
                            }

                            /**Ementa**/
                            var listaTables = Regex.Split(respostaNivel2_.ToLower(), urlTratada.ToLower().Contains("/mpv/") ? "<table border=0 width=100%" : "<table border=0").ToList<string>();

                            int index = urlTratada.ToLower().Contains("/mpv/") ? 1 : listaTables.Count - 1;

                            int indexAtual = listaTables[index].ToLower().IndexOf("<p style=");

                            int indexAtual1 = listaTables[index].ToLower().IndexOf("<p class=");

                            int indexAtual2 = listaTables[index].ToLower().IndexOf("<p align=");

                            if (indexAtual1 >= 0 && (indexAtual1 < indexAtual || indexAtual < 0) && (indexAtual1 < indexAtual2 || indexAtual2 < 0))
                                indexAtual = indexAtual1;

                            if (indexAtual2 >= 0 && (indexAtual2 < indexAtual || indexAtual < 0) && (indexAtual2 < indexAtual1 || indexAtual1 < 0))
                                indexAtual = indexAtual2;

                            int vez = 1;
                            string ementas = string.Empty;

                            listaTables.ForEach(delegate(string x) { ementas = ementas.Equals(string.Empty) && listaTables[listaTables.Count - vez].Contains("</font>") ? listaTables[listaTables.Count - vez].Substring(indexAtual).Substring(0, listaTables[listaTables.Count - vez].Substring(indexAtual).IndexOf("</font>")) : ementas; });

                            ementa.Ementa = ObterStringLimpa(ementas);

                            if (string.IsNullOrEmpty(ementa.Ementa))
                                new BuscaLegalDao().InserirLogErro(new Exception("Ementa não preenchida"), urlTratada, string.Empty);

                            /**Texto**/

                            int indexFinal = respostaNivel2_.LastIndexOf("</font>") > respostaNivel2_.LastIndexOf("</table>") ? respostaNivel2_.LastIndexOf("</font>") : respostaNivel2_.LastIndexOf("</table>");

                            ementa.Texto = @novaCss +
                                            " <p align=\"center\" style=\"margin-top: 15px; margin-bottom: 15px\"><font face=\"Arial\" color=\"#000080\" size=\"2\"><strong><a style=\"color: rgb(0,0,128)\" " +
                                            respostaNivel2_.Substring(respostaNivel2_.IndexOf("href=http://legislacao.planalto.gov.br/legisla/"), indexFinal - respostaNivel2_.IndexOf("href=http://legislacao.planalto.gov.br/legisla/"));

                            ementa.Texto = removeTagScript(ementa.Texto);

                            /**Hash**/
                            var textos = Regex.Split(respostaNivel2_, "</table>").ToList();

                            ementa.Hash = this.GerarHash(removerCaracterEspecial(ObterStringLimpa(ementa.Texto)));

                            dynamic itemUrl = new ExpandoObject();

                            itemUrl.IdUrl = itemLista_Nivel2.IdUrl;

                            itemUrl.ListaEmenta = new List<dynamic>();

                            itemUrl.ListaEmenta.Add(ementa);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemUrl };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            ex.Source = "Docs > Outros RFB > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Banco Central - Lote 3"

            #region "Captura URL's"

            dynamic objUrll;
            dynamic objItenUrl;

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                listaInserir = new List<dynamic>();

                List<string> listaUrlProcessar = new List<string>() { 
                    "http://www.bcb.gov.br/Pre/leisedecretos/Port/LeisAdmin.asp",
                    "http://www.bcb.gov.br/Pre/leisedecretos/Port/DecretosSFN.asp",
                    "http://www.bcb.gov.br/Pre/leisedecretos/Port/DecLeiSFN.asp",
                    "http://www.bcb.gov.br/Pre/leisedecretos/Port/LeisSFN.asp",
                    "http://www.bcb.gov.br/Pre/leisedecretos/Port/LeisCompl.asp"
                };

                foreach (var url in listaUrlProcessar)
                {
                    objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Banco Central";
                    objUrll.Url = url;
                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    string resposta = getHtmlPaginaByGet(url, string.Empty);

                    resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("\"", string.Empty));

                    var listaUrls = Regex.Split(resposta.Substring(resposta.IndexOf("<table")).Substring(0, resposta.Substring(resposta.IndexOf("<table")).IndexOf("</ul>")), "<li>").ToList();

                    listaUrls.RemoveAt(0);

                    listaUrls.ForEach(delegate(string x)
                    {
                        objItenUrl = new ExpandoObject();

                        objItenUrl.Url = x.Substring(x.IndexOf("href=") + 5, (x.IndexOf(".htm") + 3) - (x.IndexOf("href=") + 5));

                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    });

                    listaInserir.Add(objUrll);
                }

                new BuscaLegalDao().AtualizarFontes(listaInserir);
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("bcb"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("bcb");

                #region "Tratamento Leandro"

                //List<string> novalist;

                ///* Tratamento inicial String */

                //string arquivo = string.Empty;

                //if (!File.Exists(@"C:\Users\systax\AppData\Roaming\Skype\My Skype Received Files\NovasNCM.txt"))
                //{
                //    arquivo = LeArquivo(@"C:\Users\systax\AppData\Roaming\Skype\My Skype Received Files\Novas NCMs - TEC copy.pdf");

                //    int inexCont = 0;

                //    arquivo = arquivo.Replace("COMERCIALIZAÇÃO PROIBIDA POR TERCEIROS", string.Empty);

                //    string valor = string.Empty;
                //    int indice = 0;

                //    while (inexCont >= 0)
                //    {
                //        inexCont = arquivo.IndexOf("Nº 241, sexta-feira, 16 de dezembro de 2016");

                //        valor = string.Empty;
                //        indice = 0;

                //        if (inexCont >= 0)
                //        {
                //            while (true)
                //            {

                //                valor = arquivo.Substring(arquivo.IndexOf("Infraestrutura de Chaves Públicas Brasileira - ICP-Brasil.") + "Infraestrutura de Chaves Públicas Brasileira - ICP-Brasil.".Length + indice, 1);

                //                if (valor.Equals("1"))
                //                {
                //                    arquivo = arquivo.Remove(arquivo.IndexOf("Infraestrutura de Chaves Públicas Brasileira - ICP-Brasil.") + "Infraestrutura de Chaves Públicas Brasileira - ICP-Brasil.".Length, indice + 1);
                //                    break;
                //                }
                //                else
                //                    indice++;
                //            }

                //            arquivo = arquivo.Remove(arquivo.IndexOf("Nº 241, sexta-feira, 16 de dezembro de 2016"), ((arquivo.IndexOf("Infraestrutura de Chaves Públicas Brasileira - ICP-Brasil.") + "Infraestrutura de Chaves Públicas Brasileira - ICP-Brasil.".Length)) - arquivo.IndexOf("Nº 241, sexta-feira, 16 de dezembro de 2016"));
                //        }
                //    }

                //    File.WriteAllText(@"C:\Users\systax\AppData\Roaming\Skype\My Skype Received Files\NovasNCM.txt", arquivo);
                //}
                //else
                //{
                //    arquivo = File.ReadAllText(@"C:\Users\systax\AppData\Roaming\Skype\My Skype Received Files\NovasNCM.txt");
                //}

                ///* Fim - Tratamento inicial String  */

                ///* Identificando Itens do Grid NCM */
                //var listaNcms = Regex.Split(arquivo, "NCM DESCRIÇÃO TEC").ToList();

                //listaNcms.RemoveAt(0);

                //novalist = new List<string>();

                ///*Nivel Tabela*/
                //listaNcms.ForEach(delegate(string x)
                //{
                //    string novo1 = x.Replace("(%)", string.Empty).Substring(0, x.IndexOf("_________"));

                //    var novaLista = Regex.Split(novo1, "\n").ToList();

                //    novaLista.RemoveAt(0);
                //    novaLista.RemoveAt(novaLista.Count - 1);

                //    /*Nivel Item tabela*/

                //    string atual = string.Empty;

                //    novaLista.ForEach(delegate(string y)
                //    {
                //        if (y.Contains("6204.19.00"))
                //            atual = atual;

                //        if (!string.IsNullOrEmpty(y.Trim()))
                //        {
                //            if (((char.IsNumber(Convert.ToChar(y.Trim().Substring(0, 1))) || y.Trim().Substring(0, 1).Equals("#")) && y.Length > 10) /*Validar*/ && (string.IsNullOrEmpty(atual) || !atual.Substring(atual.Length - 1).Equals(",")))
                //            {
                //                if (!string.IsNullOrEmpty(atual))
                //                    novalist.Add(atual);

                //                atual = y;
                //            }
                //            else if (atual.Length > 1 &&
                //                        (!char.IsNumber(Convert.ToChar(atual.Substring(atual.Length - 1))) || !char.IsNumber(Convert.ToChar(y.Substring(0, 1))) || y.Length <= 2 || (y.Length <= 5 && (y.Contains("BK") || y.Contains("BIT")))) &&
                //                            !(atual.Contains("BK") || atual.Contains("BIT")))
                //            {
                //                atual += " " + y;
                //            }
                //        }
                //    });

                //    if (!string.IsNullOrEmpty(atual))
                //        novalist.Add(atual);
                //    else
                //        atual = atual;
                //});

                ///*Fim - Identificando Itens do Grid NCM*/

                ///*Quebra dos dados na Estrutura .CSV*/

                //List<string> listaFinal = new List<string>();

                //string novoNcm = string.Empty;
                //string auxNcm = string.Empty;

                //novalist.ForEach(delegate(string x)
                //{
                //    novoNcm = string.Empty;

                //    x = x.Replace(";", string.Empty).Replace("m2", "m").Replace("ASTM D 1238", "ASTM D 1238-").Replace("Capítulo 87", "Capítulo87").Replace("Capítulo 3", "Capítulo3").Replace("cm 3", "cm");

                //    /*Identificação do código*/

                //    x.ToList().ForEach(delegate(char y)
                //    {
                //        if ((char.IsNumber(y) || y.Equals(' ') || y.Equals('.')) && !novoNcm.Contains(";"))
                //            novoNcm += y.ToString();

                //        else if (novoNcm.Contains(";") || string.IsNullOrEmpty(novoNcm))
                //            novoNcm += y.ToString();

                //        else
                //            novoNcm += novoNcm.Substring(novoNcm.Length - 1).Equals(";") ? y.ToString() : ";" + y.ToString();
                //    });

                //    /*Fim - Identificação do código*/

                //    /*Nº Alíquota*/
                //    if (x.Substring(x.Length - 1).Equals("*") || x.Substring(x.Length - 1).Equals("#") || char.IsNumber(Convert.ToChar(x.Substring(x.Length - 1))) || (x.Contains("BK") || x.Contains("BIT")))
                //    {
                //        int count = 1;
                //        string aux = string.Empty;

                //        bool tentativa = false;

                //        while (true)
                //        {
                //            if ((x.Substring(x.Length - count, 1).Equals("*") || x.Substring(x.Length - count, 1).Equals("#") || char.IsNumber(Convert.ToChar(x.Substring(x.Length - count, 1)))) && (aux.Length < 2 /*Validar*/|| ((aux.Contains("*") || aux.Contains("#")) && aux.Length < 4)) && !(x.Contains("BK") || x.Contains("BIT")))
                //            {
                //                aux = x.Substring(x.Length - count, 1) + aux;
                //                count++;
                //            }
                //            else if ((x.Contains("BK") || x.Contains("BIT")) && (char.IsNumber(Convert.ToChar(x.Substring((x.IndexOf("BK") >= 0 ? x.IndexOf("BK") : x.IndexOf("BIT")) - count, 1))) || !tentativa))
                //            {
                //                aux = x.Substring((x.IndexOf("BK") >= 0 ? x.IndexOf("BK") : x.IndexOf("BIT")) - count, 1) + aux;
                //                tentativa = true;
                //                count++;
                //            }
                //            else
                //            {
                //                if (x.Substring(x.Length - count, 1).Equals(" ") || (x.Contains("BK") || x.Contains("BIT")))
                //                {
                //                    novoNcm = novoNcm.Insert(novoNcm.LastIndexOf(aux), ";");
                //                    tentativa = false;
                //                }

                //                break;
                //            }
                //        }
                //    }
                //    else
                //        novoNcm += ";;";

                //    listaFinal.Add(novoNcm);

                //    /*Fim - Nº Alíquota*/
                //});

                ///*FIM - Quebra dos dados na Estrutura .CSV*/

                ///*Criando o Arquivo*/

                //novoNcm = "Codigo_NCM;NCM;Perc(%)\n";

                //listaFinal.ForEach(x => novoNcm += x + "\n");

                //File.WriteAllText(@"C:\Users\systax\AppData\Roaming\Skype\My Skype Received Files\NovasNCM.csv", novoNcm);

                #endregion

                string urlTratada = string.Empty;

                dynamic ementa;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim();

                            itemListaVez.Url = urlTratada;

                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, "default");

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            /*Captura CSS*/
                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            string novaCss = string.Empty;

                            if (listaCss.Count > 0)
                            {

                                listaCssTratada = new List<string>();

                                listaCss.ForEach(delegate(string x)
                                {
                                    novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                    listaCssTratada.Add(novaCss);
                                });
                            }

                            /*Fim Captura CSS*/

                            ementa = new ExpandoObject();

                            /**Titulo**/
                            var titulos = respostaNivel2_.ToLower().Substring(respostaNivel2_.ToLower().IndexOf("href=http://legislacao.planalto.gov.br/legisla/"));

                            if (titulos.IndexOf("</strong>") >= 0 && (titulos.IndexOf("</strong>") < titulos.IndexOf("</small>") || titulos.IndexOf("</small>") < 0))
                                titulos = ("<" + (titulos.Substring(0, titulos.IndexOf("</strong>"))));
                            else
                                titulos = ("<" + (titulos.Substring(0, titulos.IndexOf("</small>"))));

                            ementa.TituloAto = ObterStringLimpa(titulos).Trim().ToLower();

                            /*Melhorar esse tratamento*/
                            if (urlTratada.ToLower().Contains("/decreto-lei/") || urlTratada.ToLower().Contains("/decreto/") || urlTratada.ToLower().Contains("/leis/") || urlTratada.ToLower().Contains("/lei/"))
                            {
                                /**Data**/
                                ementa.DataEdicao = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" de") + 4, (ementa.TituloAto.Length) - (ementa.TituloAto.IndexOf(" de") + 4));
                                ementa.Publicacao = ementa.DataEdicao;
                            }
                            else
                            {
                                /**Data**/
                                ementa.DataEdicao = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" de") + 3, (ementa.TituloAto.Length) - (ementa.TituloAto.IndexOf(" de") + 3));
                                ementa.Publicacao = ementa.DataEdicao;
                            }

                            try
                            {
                                /**Espécie**/
                                ementa.Especie = ementa.TituloAto.Substring(0, ementa.TituloAto.IndexOf(" n")).Trim();
                                /**Numero**/
                                //ementa.NumeroAto = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).Substring(0, ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).IndexOf(",")).Trim();
                                ementa.NumeroAto = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).Substring(0, ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).IndexOf(",") >= 0 ? ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).IndexOf(",") : ementa.TituloAto.Substring(ementa.TituloAto.IndexOf(" n") + 2).IndexOf("de")).Trim();
                            }
                            catch (Exception)
                            {
                                ementa.Especie = ementa.TituloAto.Substring(0, ementa.TituloAto.IndexOf("n")).Trim();

                                //ementa.NumeroAto = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).Substring(0, ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).IndexOf(",")).Trim();
                                ementa.NumeroAto = ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).Substring(0, ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).IndexOf(",") >= 0 ? ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).IndexOf(",") : ementa.TituloAto.Substring(ementa.TituloAto.IndexOf("n") + 2).IndexOf("de")).Trim();
                            }

                            ementa.NumeroAto = Regex.Replace(ementa.NumeroAto, "[^0-9]+", string.Empty);
                            ementa.NumeroAto = string.IsNullOrEmpty(ementa.NumeroAto) ? "0" : ementa.NumeroAto;
                            ementa.Tipo = 3;

                            /**Metadado**/
                            ementa.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                            /**Outros**/
                            ementa.Sigla = string.Empty;
                            ementa.DescSigla = string.Empty;
                            ementa.HasContent = false;

                            ementa.Republicacao = string.Empty;
                            ementa.Escopo = "FED";
                            ementa.IdFila = itemLista_Nivel2.Id;

                            /**Arquivo**/
                            ementa.ListaArquivos = new List<ArquivoUpload>();

                            if (respostaNivel2_.IndexOf("de Motivos") > 0)
                            {
                                int indexArq = 12;
                                string nomeArquivo = string.Empty;

                                while (!respostaNivel2_.Substring(respostaNivel2_.IndexOf("de Motivos") - indexArq).Substring(0, 1).Equals("="))
                                {
                                    nomeArquivo = respostaNivel2_.Substring(respostaNivel2_.IndexOf("de Motivos") - indexArq).Substring(0, 1) + nomeArquivo;
                                    indexArq++;
                                }

                                if (!nomeArquivo.Substring(nomeArquivo.LastIndexOf(".") + 1).Contains("htm"))
                                {
                                    urlTratada = urlTratada.Substring(0, urlTratada.LastIndexOf("/")).Substring(0, urlTratada.Substring(0, urlTratada.LastIndexOf("/")).LastIndexOf("/"));
                                    nomeArquivo = nomeArquivo.Substring(nomeArquivo.IndexOf("/"));

                                    string nomeArqFull = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArquivo.Substring(nomeArquivo.LastIndexOf("/") + 1));

                                    using (WebClient webClient = new WebClient())
                                    {
                                        webClient.DownloadFile(string.Format("{0}{1}", urlTratada, nomeArquivo), nomeArqFull);
                                    }

                                    byte[] arrayFile = File.ReadAllBytes(nomeArqFull);

                                    ementa.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = nomeArquivo.Substring(nomeArquivo.LastIndexOf(".") + 1), NomeArquivo = nomeArquivo.Substring(nomeArquivo.LastIndexOf("/") + 1) });

                                    File.Delete(nomeArqFull);
                                }
                            }

                            /**Ementa**/
                            var listaTables = Regex.Split(respostaNivel2_.ToLower(), urlTratada.ToLower().Contains("/mpv/") ? "<table border=0 width=100%" : "<table border=0").ToList<string>();

                            int index = urlTratada.ToLower().Contains("/mpv/") ? 1 : listaTables.Count - 1;

                            int indexAtual = listaTables[index].ToLower().IndexOf("<p style=");

                            int indexAtual1 = listaTables[index].ToLower().IndexOf("<p class=");

                            int indexAtual2 = listaTables[index].ToLower().IndexOf("<p align=");

                            if (indexAtual1 >= 0 && (indexAtual1 < indexAtual || indexAtual < 0) && (indexAtual1 < indexAtual2 || indexAtual2 < 0))
                                indexAtual = indexAtual1;

                            if (indexAtual2 >= 0 && (indexAtual2 < indexAtual || indexAtual < 0) && (indexAtual2 < indexAtual1 || indexAtual1 < 0))
                                indexAtual = indexAtual2;

                            if (indexAtual < 0 || (listaTables.Count > 3 && urlTratada.ToLower().Contains("/mpv/"))) break;

                            var ementas = listaTables[index].Substring(indexAtual).Substring(0, listaTables[index].Substring(indexAtual).IndexOf("</font>"));

                            ementa.Ementa = ObterStringLimpa(ementas);

                            if (string.IsNullOrEmpty(ementa.Ementa))
                                new BuscaLegalDao().InserirLogErro(new Exception("Ementa não preenchida"), urlTratada, string.Empty);

                            /**Texto**/
                            int indexFinal = respostaNivel2_.LastIndexOf("</font>") > respostaNivel2_.LastIndexOf("</table>") ? respostaNivel2_.LastIndexOf("</font>") : respostaNivel2_.LastIndexOf("</table>");

                            ementa.Texto = @novaCss +
                                            " <p align=\"center\" style=\"margin-top: 15px; margin-bottom: 15px\"><font face=\"Arial\" color=\"#000080\" size=\"2\"><strong><a style=\"color: rgb(0,0,128)\" " +
                                            respostaNivel2_.Substring(respostaNivel2_.IndexOf("href=http://legislacao.planalto.gov.br/legisla/"), indexFinal - respostaNivel2_.IndexOf("href=http://legislacao.planalto.gov.br/legisla/"));

                            ementa.Texto = removeTagScript(ementa.Texto);

                            /**Hash**/
                            var textos = Regex.Split(respostaNivel2_, "</table>").ToList();

                            ementa.Hash = this.GerarHash(removerCaracterEspecial(ObterStringLimpa(ementa.Texto)));

                            ementa.Republicacao = string.Empty;

                            dynamic itemUrl = new ExpandoObject();

                            itemUrl.IdUrl = itemLista_Nivel2.IdUrl;

                            itemUrl.ListaEmenta = new List<dynamic>();

                            itemUrl.ListaEmenta.Add(ementa);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemUrl };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            ex.Source = "Docs > Outros RFB > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Receita Federal(Acordo para evitar dupla tributação) - Lote 3"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                List<string> listUrl = new List<string>(){"http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/acordos-internacionais/acordos-para-evitar-a-dupla-tributacao/acordos-para-evitar-a-dupla-tributacao",
                "http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/acordos-internacionais/acordos-para-intercambio-de-informacoes-relativas-a-tributos/acordos-para-intercambio-de-informacoes-relativas-a-tributos",
                "http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/acordos-internacionais/acordos-de-complementacao-economica/acordos-de-complementacao-economica",
                "http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/acordos-internacionais/acordos-de-cooperacao-aduaneira/acordos-de-cooperacao-aduaneira"};

                listUrl.ForEach(delegate(string url)
                {
                    objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Receita Federal";
                    objUrll.Url = url;
                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    string resposta = getHtmlPaginaByGet(url, string.Empty);

                    resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("\"", string.Empty));

                    var listaUrls = Regex.Split(resposta, "<a title= href=").ToList();

                    listaUrls.RemoveRange(0, 3);

                    listaUrls.ForEach(delegate(string x)
                    {
                        objItenUrl = new ExpandoObject();

                        objItenUrl.Url = x.Substring(0, x.IndexOf("class="));

                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    });

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                });
            }
            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("rfb-edt"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("rfb-edt");

                string urlTratada = string.Empty;

                dynamic ementaInserir;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            //Todo: Tratar para salvar o arquivo PDF na base
                            if (urlTratada.Contains(".pdf"))
                                continue;

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, string.Empty);

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            /*Captura CSS*/
                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            string cssStyle = string.Empty;

                            if (listaCss.Count > 0)
                            {
                                listaCssTratada = new List<string>();

                                listaCss.ForEach(delegate(string x)
                                {
                                    string novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                    listaCssTratada.Add(novaCss);
                                });

                                listaCssTratada.ForEach(x => cssStyle += x);
                            }

                            /*Fim Captura CSS*/

                            string textoFull = string.Empty;

                            textoFull = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<div id=content>"), respostaNivel2_.IndexOf("<div id=viewlet-below-content>") - respostaNivel2_.IndexOf("<div id=content>"));

                            string republicacao = ObterStringLimpa("<" + textoFull.Substring(textoFull.IndexOf("class=documentPublished") + 6).Substring(0, textoFull.Substring(textoFull.IndexOf("class=documentPublished") + 6).IndexOf("class=")));

                            string modificacao = ObterStringLimpa("<" + textoFull.Substring(textoFull.IndexOf("class=documentModified")).Substring(0, textoFull.IndexOf("</div>")));

                            string ementaTexto = ObterStringLimpa(textoFull.Substring(textoFull.IndexOf("<table")).Substring(0, textoFull.Substring(textoFull.IndexOf("<table")).IndexOf("</table>")));

                            int inicioCorpo = textoFull.IndexOf("<h1 class=TituloPaginas>");
                            string publicacao = string.Empty;

                            if (inicioCorpo < 0)
                            {
                                inicioCorpo = textoFull.IndexOf("<h1 class=documentFirstHeading>");
                                publicacao = ObterStringLimpa(textoFull.Substring(textoFull.IndexOf("<div id=content-core>")).Substring(textoFull.Substring(textoFull.IndexOf("<div id=content-core>")).IndexOf("<p"), textoFull.Substring(textoFull.IndexOf("<div id=content-core>")).IndexOf("</p>") - textoFull.Substring(textoFull.IndexOf("<div id=content-core>")).IndexOf("<p")));
                            }
                            else
                            {
                                publicacao = ObterStringLimpa(textoFull.Substring(textoFull.IndexOf("</table>")).Substring(0, textoFull.Substring(textoFull.IndexOf("</table>")).IndexOf("</p>")));

                                if (publicacao.Length > 50 || !publicacao.Contains("DOU"))
                                {
                                    publicacao = ObterStringLimpa(textoFull.Substring(textoFull.LastIndexOf("</h1>")).Substring(0, textoFull.Substring(textoFull.LastIndexOf("</h1>")).IndexOf("</b>") < 0 ? 0 : textoFull.Substring(textoFull.LastIndexOf("</h1>")).IndexOf("</b>")));
                                }
                            }

                            if ((publicacao.Length > 50 && !publicacao.Contains("DOU")) || !publicacao.Contains("DOU"))
                            {
                                publicacao = ObterStringLimpa(textoFull.Substring(textoFull.IndexOf("<div id=content-core>")).Substring(textoFull.Substring(textoFull.IndexOf("<div id=content-core>")).IndexOf("<p"), textoFull.Substring(textoFull.IndexOf("<div id=content-core>")).IndexOf("</p>") - textoFull.Substring(textoFull.IndexOf("<div id=content-core>")).IndexOf("<p")));

                                if (publicacao.Length > 50 || !publicacao.Contains("DOU"))
                                    publicacao = string.Empty;
                            }

                            string titulo = ObterStringLimpa(textoFull.Substring(inicioCorpo).Substring(0, textoFull.Substring(inicioCorpo).IndexOf("</h1>")));

                            int indexQuebra = titulo.ToLower().IndexOf("nº") < 0 ?
                                                titulo.ToLower().IndexOf("n°") < 0 ?
                                                    titulo.ToLower().IndexOf("n °") < 0 ?
                                                        titulo.ToLower().IndexOf("n.º") < 0 ?
                                                            titulo.ToLower().IndexOf("n º") : titulo.ToLower().IndexOf("n.º") : titulo.ToLower().IndexOf("n °") : titulo.ToLower().IndexOf("n°") : titulo.ToLower().IndexOf("nº");

                            string especie = string.Empty;
                            string numero = string.Empty;
                            string edicao = string.Empty;

                            if (indexQuebra >= 0)
                            {
                                especie = titulo.Substring(0, indexQuebra);
                                numero = titulo.Trim().Substring(indexQuebra + 2, (titulo.ToLower().IndexOf(",") < 0 ? titulo.Trim().ToLower().IndexOf(" de") : titulo.ToLower().IndexOf(",")) - (indexQuebra + 2));
                                edicao = titulo.Trim().Substring((titulo.Trim().ToLower().IndexOf(",") < 0 ? titulo.Trim().ToLower().IndexOf(" de") : titulo.Trim().ToLower().IndexOf(","))); ;
                            }

                            ementaInserir = new ExpandoObject();

                            /**Outros**/
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /**Arquivo**/
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                            /**Default**/
                            ementaInserir.Publicacao = publicacao.Trim();
                            ementaInserir.Republicacao = republicacao.Replace("publicado", string.Empty).Replace(",", string.Empty).Trim();
                            ementaInserir.Ementa = ementaTexto.Trim();
                            ementaInserir.TituloAto = titulo.Trim();
                            ementaInserir.Especie = especie.Trim();
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao.Replace(", de", string.Empty).Replace(" de", string.Empty).Trim();
                            ementaInserir.Texto = cssStyle + Regex.Replace(removeTagScript(textoFull.Trim()), @"\s+", " ");
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(textoFull)));

                            ementaInserir.Escopo = "FED";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Procuradoria Geral da Fazenda Nacional - Lote 3"

            #region "Captura URL's"
            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                listaInserir = new List<dynamic>();

                List<string> listaUrlProcessar = new List<string>() { "http://dados.pgfn.fazenda.gov.br/dataset/pareceres", "http://dados.pgfn.fazenda.gov.br/dataset/notas" };

                foreach (var url in listaUrlProcessar)
                {
                    int countPage = 1;

                    while (true)
                    {
                        try
                        {
                            objUrll = new ExpandoObject();

                            objUrll.Indexacao = "Procuradoria Geral da Fazenda Nacional";
                            objUrll.Url = url;
                            objUrll.Lista_Nivel2 = new List<dynamic>();

                            string resposta = getHtmlPaginaByGet(url + "?page=" + countPage.ToString(), string.Empty);

                            resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("\"", string.Empty));

                            if (resposta.Contains("(Nenhum) "))
                                break;

                            var listaUrls = Regex.Split(resposta.Substring(resposta.IndexOf("class=main-link")).Substring(0, resposta.Substring(resposta.IndexOf("class=main-link")).LastIndexOf("class=pagination")), "class=main-link").ToList();

                            if (string.IsNullOrEmpty(listaUrls[0]))
                                listaUrls.RemoveAt(0);

                            listaUrls.ForEach(delegate(string x)
                            {
                                objItenUrl = new ExpandoObject();

                                string urlNova = x.Substring(x.IndexOf("=") + 1, x.IndexOf(">") - (x.IndexOf("=") + 1));

                                objItenUrl.Url = url + urlNova.Substring(urlNova.IndexOf("/resource"));

                                objUrll.Lista_Nivel2.Add(objItenUrl);
                            });

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });

                            countPage++;
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
            }
            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("pgfn"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("pgfn");

                string urlTratada = string.Empty;

                dynamic ementaInserir;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim();

                            itemListaVez.Url = urlTratada;

                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, string.Empty);

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            /*Captura CSS*/

                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            string xx = string.Empty;

                            if (listaCss.Count > 0)
                            {

                                listaCssTratada = new List<string>();

                                listaCss.ForEach(delegate(string x)
                                {
                                    string novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                    listaCssTratada.Add(novaCss);
                                });

                                listaCssTratada.ForEach(x => xx += x);
                            }

                            /*Fim Captura CSS*/

                            string textoFull = string.Empty;

                            textoFull = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<h1 class=page_heading>"), respostaNivel2_.IndexOf("<a href=#top>Voltar ao topo</a>") - respostaNivel2_.IndexOf("<h1 class=page_heading>"));

                            string titulo = ObterStringLimpa(textoFull.Substring(textoFull.IndexOf("<h1"), textoFull.IndexOf("</h1>") - textoFull.IndexOf("<h1")));

                            string numeroAto = titulo.Substring(titulo.ToLower().IndexOf("nº") + 2, titulo.IndexOf("/") - (titulo.ToLower().IndexOf("nº") + 2));

                            string ementa = ObterStringLimpa(textoFull.Substring(textoFull.IndexOf("<h2")).Substring(0, textoFull.Substring(textoFull.IndexOf("<h2")).IndexOf("</div>")));

                            /** Arquivo **/
                            string arquivo = textoFull.Substring(textoFull.IndexOf("class=pretty-button primary resource-url-analytics") + "class=pretty-button primary resource-url-analytics".Length);
                            arquivo = arquivo.Substring(arquivo.IndexOf("href=") + 5, arquivo.IndexOf(">") - (arquivo.IndexOf("href=") + 5));

                            string nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), arquivo.Substring(arquivo.LastIndexOf("/") + 1));

                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(arquivo, nomeArq);
                            }

                            byte[] arrayFile = File.ReadAllBytes(nomeArq);

                            /** Fim-Arquivo **/

                            File.Delete(nomeArq);

                            textoFull = removeTagScript(textoFull);

                            string hash = GerarHash(removerCaracterEspecial((textoFull.Substring(textoFull.LastIndexOf("<h3")))));

                            var listDados = Regex.Split(textoFull.Substring(textoFull.IndexOf("<h3")).Substring(0, textoFull.Substring(textoFull.IndexOf("<h3")).IndexOf("</table>")), "<td class").ToList();

                            listDados.RemoveRange(0, 2);

                            string sigla = ObterStringLimpa("<" + listDados[0]).Trim();
                            string data = ObterStringLimpa("<" + listDados[listDados.Count - 1]).Trim();

                            ementaInserir = new ExpandoObject();

                            /**Outros**/
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.DescSigla = sigla;
                            ementaInserir.HasContent = false;

                            /**Arquivo**/
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();
                            ementaInserir.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = nomeArq.Substring(nomeArq.LastIndexOf(".") + 1), NomeArquivo = arquivo.Substring(arquivo.LastIndexOf("/") + 1) });

                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                            /**Default**/
                            ementaInserir.Publicacao = data.Trim();
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo.Trim();
                            ementaInserir.Especie = ementaInserir.TituloAto.Substring(0, ementaInserir.TituloAto.IndexOf(" "));
                            ementaInserir.NumeroAto = Regex.Replace(numeroAto, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = xx + textoFull.Trim();
                            ementaInserir.Hash = hash;

                            ementaInserir.Escopo = "FED";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Decisões Vinculantes do STF e do STJ (repercussão geral e recursos repetitivos) - Lote 3"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                listaInserir = new List<dynamic>();

                List<string> listaUrlProcessar = new List<string>() { "http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/decisoes-vinculantes-do-stf-e-do-stj-repercussao-geral-e-recursos-repetitivos" };

                foreach (var url in listaUrlProcessar)
                {
                    objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Receita Federal";
                    objUrll.Url = url;
                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    string resposta = getHtmlPaginaByGet(url, string.Empty);

                    resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("\"", " "));

                    var parte1 = resposta.Substring(resposta.IndexOf("</em>")).Substring(0, resposta.Substring(resposta.IndexOf("</em>")).IndexOf("</ul>"));

                    var parte2 = resposta.Substring(resposta.IndexOf("</em>")).Substring(resposta.Substring(resposta.IndexOf("</em>")).IndexOf("</ul>") + 3);

                    parte2 = parte2.Substring(parte2.IndexOf("<ul"), parte2.IndexOf("</ul>") - parte2.IndexOf("<ul"));

                    var listaUrls = Regex.Split(parte1 + parte2, "<li").ToList();

                    listaUrls.RemoveAt(0);

                    listaUrls.ForEach(delegate(string x)
                    {
                        objItenUrl = new ExpandoObject();

                        string urlNova = x.Substring(x.IndexOf("href=") + 5).Substring(0, x.Substring(x.IndexOf("href=") + 5).IndexOf(" "));

                        urlNova += string.Format("|{0}", x.Substring(x.IndexOf("href=") + 5).Substring(x.Substring(x.IndexOf("href=") + 5).IndexOf(">") + 1, x.Substring(x.IndexOf("href=") + 5).IndexOf("</a>") - (x.Substring(x.IndexOf("href=") + 5).IndexOf(">") + 1)));

                        objItenUrl.Url = urlNova;

                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    });

                    listaInserir.Add(objUrll);
                }

                new BuscaLegalDao().AtualizarFontes(listaInserir);
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("rfb-dvss"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("rfb-dvss");

                string urlTratada = string.Empty;

                dynamic ementaInserir;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            /** Arquivo **/
                            string nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), urlTratada.Split('|')[0].Substring(urlTratada.Split('|')[0].LastIndexOf("/") + 1));

                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(urlTratada.Split('|')[0], nomeArq);
                            }

                            byte[] arrayFile = File.ReadAllBytes(nomeArq);
                            string conteudoPdf = LeArquivo(nomeArq);
                            File.Delete(nomeArq);

                            /** Fim-Arquivo **/
                            string hash = GerarHash(removerCaracterEspecial(Regex.Replace(conteudoPdf.Replace("\n", " ").Replace("\t", " ").Trim(), @"\s+", " ")));

                            ementaInserir = new ExpandoObject();

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();
                            ementaInserir.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = nomeArq.Substring(nomeArq.LastIndexOf(".") + 1), NomeArquivo = urlTratada.Split('|')[0].Substring(urlTratada.Split('|')[0].LastIndexOf("/") + 1) });

                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                            var dadosItem = urlTratada.Split('|')[1].Split(' ');

                            /** Default **/
                            ementaInserir.Publicacao = dadosItem[dadosItem.Length - 1].Substring(dadosItem[dadosItem.Length - 1].IndexOf("/") + 1);
                            ementaInserir.Sigla = dadosItem[1];
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = string.Empty;
                            ementaInserir.TituloAto = urlTratada.Split('|')[1].Trim();
                            ementaInserir.Especie = dadosItem[0];
                            ementaInserir.NumeroAto = Regex.Replace(dadosItem[dadosItem.Length - 1].Substring(0, dadosItem[dadosItem.Length - 1].IndexOf("/")), "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = Regex.Replace(conteudoPdf.Trim(), @"\s+", " ");
                            ementaInserir.Hash = hash;

                            ementaInserir.Escopo = "FED";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Banco Central - Normativos - Lote 3"

            #region "Captura URL's"

            List<string> urlPronta;

            string tipo;
            string dataa;
            string numeroo;

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                List<string> listaUrls = new List<string>() {"http://www.bcb.gov.br/pre/normativos/busca/buscaSharePoint.asp?tipo=Ato%20do%20Presidente&startRow={0}",
                "http://www.bcb.gov.br/pre/normativos/busca/buscaSharePoint.asp?tipo=Ato%20de%20Diretor&startRow={0}",
                "http://www.bcb.gov.br/pre/normativos/busca/buscaSharePoint.asp?tipo=Carta%20Circular&startRow={0}",
                "http://www.bcb.gov.br/pre/normativos/busca/buscaSharePoint.asp?tipo=Circular&startRow={0}",
                "http://www.bcb.gov.br/pre/normativos/busca/buscaSharePoint.asp?tipo=Comunicado&startRow={0}",
                "http://www.bcb.gov.br/pre/normativos/busca/buscaSharePoint.asp?tipo=Comunicado%20Conjunto&startRow={0}",
                "http://www.bcb.gov.br/pre/normativos/busca/buscaSharePoint.asp?tipo=Decis%C3%A3o%20Conjunta&startRow={0}",
                "http://www.bcb.gov.br/pre/normativos/busca/buscaSharePoint.asp?tipo=Ato%20Normativo%20Conjunto&startRow={0}",
                "http://www.bcb.gov.br/pre/normativos/busca/buscaSharePoint.asp?tipo=Resolu%C3%A7%C3%A3o&startRow={0}"};

                int indexVez = 0;

                tipo = string.Empty;
                dataa = string.Empty;
                numeroo = string.Empty;

                foreach (var url in listaUrls)
                {
                    /*Remover depois*/
                    urlPronta = new List<string>();

                    listaInserir = new List<dynamic>();

                    while (true)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            WebRequest req = WebRequest.Create(string.Format(url, indexVez.ToString()));

                            req.Headers.Add(HttpRequestHeader.AcceptLanguage, "application/json; odata=verbose");
                            req.Headers.Add(HttpRequestHeader.AcceptEncoding, "application/json; odata=verbose");

                            WebResponse res = req.GetResponse();
                            Stream dataStream = res.GetResponseStream();
                            StreamReader reader = new StreamReader(dataStream);

                            objUrll = new ExpandoObject();

                            objUrll.Indexacao = "Banco Central";
                            objUrll.Url = url;
                            objUrll.Lista_Nivel2 = new List<dynamic>();

                            string resposta = reader.ReadToEnd();

                            resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("\"", " "));

                            if (!resposta.Contains("Key:title,Value:"))
                                break;

                            var listUrl = Regex.Split(resposta, "Key:title,Value:").ToList();

                            listUrl.RemoveAt(0);

                            listUrl.ForEach(delegate(string x)
                            {
                                tipo = url.Contains("Ato%20Normativo%20Conjunto") ? "Ato%20Normativo%20Conjunto" :
                                            url.Contains("Decis%C3%A3o%20Conjunta") ? "Decis%C3%A3o%20Conjunta" :
                                                url.Contains("Ato%20do%20Presidente") ? "Ato%20do%20Presidente" :
                                                    url.Contains("Ato%20de%20Diretor") ? "Ato%20de%20Diretor" :
                                                        url.Contains("Carta%20Circular") ? "Carta%20Circular" :
                                                            url.Contains("=Circular") ? "Circular" :
                                                                url.Contains("Comunicado%20Conjunto") ? "Comunicado%20Conjunto" :
                                                                    url.Contains("Comunicado&") ? "Comunicado" : "Resolu%C3%A7%C3%A3o";
                                //x.Substring(0, (x.IndexOf(@"N\") < 0 ? x.IndexOf(@"N.") : x.IndexOf(@"N\")));

                                dataa = x.Substring(x.IndexOf("Key:RefinableString01,Value:string;#") + "Key:RefinableString01,Value:string;#".Length).Substring(0, x.Substring(x.IndexOf("Key:RefinableString01,Value:string;#") + "Key:RefinableString01,Value:string;#".Length).IndexOf(" "));
                                numeroo = x.Substring(x.IndexOf("Key:NumeroOWSNMBR,Value:") + "Key:NumeroOWSNMBR,Value:".Length).Substring(0, x.Substring(x.IndexOf("Key:NumeroOWSNMBR,Value:") + "Key:NumeroOWSNMBR,Value:".Length).IndexOf("."));

                                /*Remover depois*/
                                urlPronta.Add(string.Format("http://www.bcb.gov.br/pre/normativos/busca/normativo.asp?numero={0}&tipo={1}&data={2}", numeroo.Trim(), tipo.Replace("-", "%20").Trim(), dataa.Trim()));

                                /*Agrupando as url para inserir todas juntas por tipo do Documento as url sao agrupadas para que possam ser vinculadas aos seus links de junção*/
                                objItenUrl = new ExpandoObject();

                                objItenUrl.Url = string.Format("http://www.bcb.gov.br/pre/normativos/busca/normativo.asp?numero={0}&tipo={1}&data={2}", numeroo.Trim(), tipo.Replace("-", "%20").Trim(), dataa.Trim());

                                objUrll.Lista_Nivel2.Add(objItenUrl);

                                listaInserir.Add(objUrll);
                            });

                            indexVez = 15 + indexVez;
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }

                    new BuscaLegalDao().AtualizarFontes(listaInserir);
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("bcb-normas"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("bcb-normas");

                string urlTratada = string.Empty;

                dynamic ementaInserir;

                /*Captura css*/

                string urlCss = "http://www.bcb.gov.br/pre/normativos/busca/normativo.asp?numero=616&tipo=Ato%20de%20Diretor&data=23/12/2016";

                string _respostaNivel = getHtmlPaginaByGet(urlCss, string.Empty);

                var listaCss = Regex.Split(_respostaNivel.Replace("\"", string.Empty).Replace("\n", string.Empty), "<link").ToList();

                listaCss.RemoveAt(0);

                string cssDoc = string.Empty;

                if (listaCss.Count > 0)
                {

                    listaCssTratada = new List<string>();

                    listaCss.ForEach(delegate(string x)
                    {
                        string novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                        novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlCss.Substring(urlCss.IndexOf("//") + 2).Substring(0, urlCss.Substring(urlCss.IndexOf("//") + 2).IndexOf("/"))));

                        listaCssTratada.Add(novaCss);
                    });

                    listaCssTratada.ForEach(x => cssDoc += x);
                }

                /*Fim Captura css*/

                int marcaInicio = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (marcaInicio < 0)
                            {
                                marcaInicio++;
                                continue;
                            }

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim();

                            itemListaVez.Url = urlTratada;

                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = urlTratada.Substring(urlTratada.IndexOf("numero=") + 7, urlTratada.IndexOf("&") - (urlTratada.IndexOf("numero=") + 7));
                            string data1 = urlTratada.Substring(urlTratada.IndexOf("data=") + 5);
                            string tipo_ = urlTratada.Substring(urlTratada.IndexOf("tipo=") + 5).Substring(0, urlTratada.Substring(urlTratada.IndexOf("tipo=") + 5).IndexOf("&"));

                            if (urlTratada.Contains("tipo=Carta%20Circular") || urlTratada.Contains("tipo=Circular"))
                                urlTratada = string.Format("http://www.bcb.gov.br/pre/normativos/busca/sharepointproxyDetalharNormativo.asp?metodo=Normativos&numero={0}&tipo={1}", numero, tipo_);
                            else
                                urlTratada = string.Format("http://www.bcb.gov.br/pre/normativos/busca/sharepointproxyDetalharNormativo.asp?metodo=Demais%20Normativos&numero={0}&tipo={1}", numero, tipo_);

                            WebRequest reqNivel2_1 = WebRequest.Create(urlTratada);

                            reqNivel2_1.Headers.Add(HttpRequestHeader.AcceptLanguage, "pt-BR,pt;q=0.8,en-US;q=0.6,en;q=0.4");
                            reqNivel2_1.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip, deflate, sdch");
                            reqNivel2_1.Headers.Add(HttpRequestHeader.Upgrade, "1");
                            reqNivel2_1.Headers.Add("X-Requested-With", "XMLHttpRequest");

                            WebResponse resNivel2_1 = reqNivel2_1.GetResponse();
                            Stream dataStreamNivel2_1 = resNivel2_1.GetResponseStream();
                            StreamReader readerNivel2_1;

                            readerNivel2_1 = new StreamReader(dataStreamNivel2_1);

                            string respostaNivel2_1 = readerNivel2_1.ReadToEnd();

                            string ementa = string.Empty;
                            string titulo_ = string.Empty;
                            string texto = string.Empty;
                            string textoJson = string.Empty;

                            if (urlTratada.Contains("tipo=Carta%20Circular") || urlTratada.Contains("tipo=Circular"))
                            {
                                var objDados = JObject.Parse(respostaNivel2_1);

                                List<string> conteudoArq = new List<string>();

                                Regex.Split(objDados["d"]["results"][0]["DocumentosAnexados"].ToString(), "#;").ToList().ForEach(delegate(string item)
                                {
                                    /** Arquivo **/
                                    if (!item.Trim().Equals(string.Empty))
                                    {
                                        string nomeArq = item.Substring(0, item.IndexOf(";"));
                                        string pathSaveArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                                        nomeArq = "http://www.bcb.gov.br/pre/normativos/busca/downloadNormativo.asp?arquivo=/Lists/Normativos/Attachments/" + objDados["d"]["results"][0]["ID"] + "/" + nomeArq;

                                        using (WebClient webClient = new WebClient())
                                        {
                                            webClient.DownloadFile(nomeArq, pathSaveArq);
                                        }

                                        byte[] arrayFile = File.ReadAllBytes(pathSaveArq);
                                        string conteudoPdf = LeArquivo(pathSaveArq);
                                        File.Delete(pathSaveArq);
                                        /** Fim-Arquivo **/

                                        conteudoArq.Add((nomeArq.Substring(nomeArq.Length - 5, 1).Equals("O") ? "O" : "L") + conteudoPdf);
                                    }
                                });

                                string dados = string.Empty;

                                numero = objDados["d"]["results"][0]["Numero"].ToString();
                                data1 = objDados["d"]["results"][0]["DataTexto"].ToString().Replace("00:00", string.Empty);
                                titulo_ = objDados["d"]["results"][0]["Title"].ToString() + ", de " + data1;
                                ementa = objDados["d"]["results"][0]["Assunto_x0020_Normativo"].ToString();
                                tipo_ = objDados["d"]["results"][0]["Tipo_x0020_do_x0020_Normativo"].ToString();

                                texto = "<h1>" + objDados["d"]["results"][0]["Title"] + (objDados["d"]["results"][0]["Revogado"].ToString().Trim().Equals("True") ? "(REVOGADO)" : string.Empty) + "</h1><div id='conteudo'><br><br>";

                                texto += @"<div class='fundoPadraoAClaro1 txN'>DOU</div>
                                <p class='justificado' id='assunto' style='margin-left:10px'><br></p><span>
                                <img src='http://www.bcb.gov.br/img/adver_12x12.gif' alt='Advertência'>
                                <span style='color: rgb(255, 0, 0);'>" + (objDados["d"]["results"][0]["DOU"].ToString().Trim().Equals(string.Empty) ? "Os textos não substituem a publicação no DOU e no Sisbacen.<br></span><br>" : objDados["d"]["results"][0]["DOU"].ToString().Trim());

                                texto += @"<div class='fundoPadraoAClaro1 txN'>Assunto</div>
                                <p class='justificado' id='assunto' style='margin-left:10px'>
                                <br>" + objDados["d"]["results"][0]["Assunto_x0020_Normativo"] + "</p><br>";

                                texto += @"<div class='fundoPadraoAClaro1 txN'>Normas vinculadas</div> <p></p>";

                                string itensNormativos = string.Empty;

                                Regex.Split(objDados["d"]["results"][0]["Normativos_x0020_Vinculados"].ToString(), ";#").ToList().ForEach(delegate(string item)
                                {
                                    if (!item.Trim().Equals(string.Empty))
                                    {
                                        var listDados = Regex.Split(item.Replace(":", ";"), ";@").ToList();

                                        if (itensNormativos.Contains(listDados[0]))
                                        {
                                            int indexRaiz = itensNormativos.IndexOf(listDados[0]);

                                            indexRaiz += itensNormativos.Substring(indexRaiz).IndexOf("</span>");

                                            itensNormativos = itensNormativos.Insert(indexRaiz, "<a href='#' title=''>" + listDados[1] + "/" + listDados[2] + "</a> |");
                                        }
                                        else
                                            itensNormativos += "<ul class='lista2' style='margin-left:10px'><li title='item'>" + listDados[0] + "<br><span id='resolucoes'><a href='#' title=''>" + listDados[1] + "/" + listDados[2] + "</a> | </span></li></ul>";
                                    }
                                });

                                texto += itensNormativos + "<p></p><br>";

                                texto += @"<div class='fundoPadraoAClaro1 txN'>Referências</div>
                                <p class='texto_comentario' style='padding-left:10px'><br>Base Legal e Regulamentar, Citações e Revogações</p><p></p><p></p>
                                <ul class='lista2' id='referencias' style='margin-left:10px'>";

                                Regex.Split(objDados["d"]["results"][0]["Referencias"].ToString(), ";#").ToList().ForEach(delegate(string item)
                                {
                                    if (!item.Trim().Equals(string.Empty))
                                        texto += "<li class='item'>" + item + "</li><br>";
                                });

                                texto += "</ul><p></p><br>";

                                texto += @"<div class='fundoPadraoAClaro1 txN'>Atualizações</div><p></p>
                                           <ul class='lista2' id='atualizacoes' style='margin-left:10px'>";

                                Regex.Split(objDados["d"]["results"][0]["Atualizacoes"].ToString(), ";#").ToList().ForEach(delegate(string item)
                                {
                                    if (!item.Trim().Equals(string.Empty))
                                        texto += "<li class='item'>" + item + "</li><br>";
                                });

                                texto += "</ul></span></div>";

                                texto += @"<div class='fundoPadraoAClaro1 txN'>Texto Vigente</div>
                                           <p class='justificado' id='assunto' style='margin-left:10px'>
                                           <br>" + conteudoArq.Find(x => !x.Substring(0, 1).Equals("O")) + "</p><br>";

                                texto += @"<div class='fundoPadraoAClaro1 txN'>Texto Original</div>
                                           <p class='justificado' id='assunto' style='margin-left:10px'>
                                           <br>" + conteudoArq.Find(x => x.Substring(0, 1).Equals("O")) + "</p><br>";

                                textoJson = Regex.Replace(texto.Trim(), @"\s+", " ");
                            }

                            else
                            {
                                respostaNivel2_1 = JObject.Parse(respostaNivel2_1).ToString();

                                respostaNivel2_1 = System.Net.WebUtility.UrlDecode(respostaNivel2_1.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                                /*Tratamento Ementa*/
                                var indexDiv = respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Title:")).IndexOf("<div");
                                var indexTexto = respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Title:")).IndexOf("Texto:");

                                if (indexDiv > indexTexto)
                                    ementa = respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Title:")).Substring(0, respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Title:")).IndexOf(","));
                                else
                                    ementa = respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Title:")).Substring(respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Title:")).IndexOf("<div"), respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Title:")).IndexOf("</div>") - respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Title:")).IndexOf("<div"));

                                /*Fim - Tratamento Ementa*/

                                textoJson = respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Texto:") + 6).Substring(0, respostaNivel2_1.Substring(respostaNivel2_1.IndexOf("Texto:") + 6).IndexOf("    Responsavel:"));

                                textoJson = Regex.Replace(textoJson.Trim(), @"\s+", " ");

                                titulo_ = string.Format("{0} nº{1}, de {2}", System.Net.WebUtility.UrlDecode(tipo_), numero, data1);

                                texto = @"<style>#conteudo {text-indent:0cm;word-wrap:break-word;}</style>
                                                 <style type=text/css>@charset 'UTF-8';[ng\:cloak],[ng-cloak],[data-ng-cloak],[x-ng-cloak],.ng-cloak,.x-ng-cloak,.ng-hide:not(.ng-hide-animate){display:none !important;}ng\:form{display:block;}</style>";

                                texto += string.Format(@"<h1>{0}</h1><div id='conteudo'><br><br><p align='' style='text-indent: 2.5cm;'></p><hr><br>
                                                            <div style='text-align:center;font-size:12px;font-family:courier new'>{1}</div><br><br>
                                                            <div style='padding-left:47%;padding-right:17.9%;text-align:justify;text-indent:0px;font-size:12px;font-family:courier new'>{2}</div><br>
                                                            <br><div style='padding-left:17.9%;padding-right:17.9%;text-align:justify;text-indent:0px;font-size:12px;font-family:courier new'>{3}</div></div>", titulo_, titulo_, ementa, textoJson);

                                ementa = indexDiv > indexTexto ? string.Empty : Regex.Replace(ObterStringLimpa(ementa.Replace("\n", " ").Replace("\t", " ")), @"\s+", " ");
                            }

                            /**Tratamento para inserir na base**/
                            ementaInserir = new ExpandoObject();

                            /**Outros**/
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /**Arquivo**/
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                            /**Default**/
                            ementaInserir.Publicacao = data1;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo_;
                            ementaInserir.Especie = System.Net.WebUtility.UrlDecode(tipo_);
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = cssDoc + texto;
                            ementaInserir.IdFila = itemLista_Nivel2.Id;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(textoJson.Replace("\n", " ").Replace("\t", " "))));

                            ementaInserir.Escopo = "FED";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            ex.Source = "Docs > BCB-NORMAS > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Resoluções CGSIM - PDF - Lote 3"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                objUrll = new ExpandoObject();

                objUrll.Indexacao = "Receita Federal";
                objUrll.Url = "http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/atos-cgsim/resolucoes-cgsim.pdf";
                objUrll.Lista_Nivel2 = new List<dynamic>();

                objItenUrl = new ExpandoObject();

                objItenUrl.Url = "http://idg.receita.fazenda.gov.br/acesso-rapido/legislacao/atos-cgsim/resolucoes-cgsim.pdf";
                objUrll.Lista_Nivel2.Add(objItenUrl);

                new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("r-cgsim"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("r-cgsim");

                /** Arquivo **/
                string nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), "resolucoes-cgsim.pdf");

                using (WebClient webClient = new WebClient())
                {
                    webClient.DownloadFile(listaUrl[0].Lista_Nivel2[0].Url, nomeArq);
                }

                string conteudoPdf = LeArquivo(nomeArq);
                File.Delete(nomeArq);

                var listResolucoes = Regex.Split(conteudoPdf.Replace("\n", string.Empty), "Presidente do").ToList();

                var listaPdf = new List<string>();

                string atual = string.Empty;

                string novoItem = string.Empty;

                //Obtendo os Dados Separados do PDF.
                listResolucoes.ForEach(delegate(string x)
                {
                    if (x.Substring(0, x.IndexOf(",")).ToLower().Contains(" nº"))
                    {
                        if (x.Contains("Resolução CGSIM nº 3, de 1º de julho de 2009"))
                        {
                            novoItem = x.Substring(x.IndexOf("Resolução CGSIM nº 2, de 1º de julho de 2009"));
                            x = x.Substring(0, x.IndexOf("Resolução CGSIM nº 2, de 1º de julho de 2009"));
                            listaPdf.Add(novoItem);
                        }

                        else

                            if (!atual.Equals(string.Empty))
                                listaPdf.Add(atual.Replace("_________________________", string.Empty));

                        atual = x;
                    }

                    else
                    {
                        if (x.Contains("_________________________  Resolução CGSIM nº 10, de 7 de outubro de 2009"))
                        {
                            novoItem = x.Substring(x.IndexOf("_________________________  Resolução CGSIM nº 10, de 7 de outubro de 2009"));
                            atual += x.Substring(0, x.IndexOf("_________________________  Resolução CGSIM nº 10, de 7 de outubro de 2009"));
                            listaPdf.Add(novoItem.Replace("_________________________", string.Empty));
                        }

                        else if (x.Contains("Resolução CGSIM nº 15, de 17 de dezembro de 2009"))
                        {
                            novoItem = x.Substring(x.IndexOf("Resolução CGSIM nº 15, de 17 de dezembro de 2009"));
                            atual += x.Substring(0, x.IndexOf("Resolução CGSIM nº 15, de 17 de dezembro de 2009"));
                            listaPdf.Add(novoItem.Replace("_________________________", string.Empty));
                        }

                        else if (x.Contains("Resolução CGSIM nº 21, de 9 de junho de 2010") && !x.Contains("Incluído pela Resolução CGSIM nº 21, de 9 de junho de 2010"))
                        {
                            novoItem = x.Substring(x.IndexOf("Resolução CGSIM nº 21, de 9 de junho de 2010"));
                            atual += x.Substring(0, x.IndexOf("Resolução CGSIM nº 21, de 9 de junho de 2010"));
                            listaPdf.Add(novoItem.Replace("_________________________", string.Empty));
                        }

                        else if (x.Contains("Resolução CGSIM nº 23, de 21 de setembro de 2010"))
                        {
                            novoItem = x.Substring(x.IndexOf("Resolução CGSIM nº 23, de 21 de setembro de 2010"));
                            atual += x.Substring(0, x.IndexOf("Resolução CGSIM nº 23, de 21 de setembro de 2010"));
                            listaPdf.Add(novoItem.Replace("_________________________", string.Empty));
                        }
                        else
                            atual += x;
                    }
                });

                listaPdf.Add(atual.Replace("_________________________", string.Empty));

                //Montando os objetos Separados do PDF para inserir na base.
                listaPdf.ForEach(delegate(string item)
                {
                    item = item.Substring(item.ToLower().IndexOf("resolução"));

                    string titulo = item.Substring(0, item.IndexOf(","));

                    string publicacao = item.IndexOf(" DOU") < item.IndexOf(". ") && item.IndexOf(" DOU") > 0 ?
                                            item.Substring(item.IndexOf(" DOU"), 18) : item.Substring(item.IndexOf(",") + 1).Substring(0, item.Substring(item.IndexOf(",") + 1).IndexOf(". ")).Replace(" DE", string.Empty);

                    string numero = titulo.Substring(titulo.ToLower().IndexOf("nº") + 2);
                    string especie = titulo.Substring(0, titulo.ToLower().IndexOf("nº"));

                    dynamic itemListaVez = new ExpandoObject();

                    itemListaVez.ListaEmenta = new List<dynamic>();
                    itemListaVez.Url = listaUrl[0].Lista_Nivel2[0].Url;
                    itemListaVez.IdUrl = listaUrl[0].Lista_Nivel2[0].IdUrl;

                    /**Tratamento para inserir na base**/
                    dynamic ementaInserir = new ExpandoObject();

                    /**Outros**/
                    ementaInserir.Sigla = string.Empty;
                    ementaInserir.DescSigla = string.Empty;
                    ementaInserir.HasContent = false;

                    /**Arquivo**/
                    ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                    ementaInserir.Tipo = 3;
                    ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                    /**Default**/
                    ementaInserir.Publicacao = publicacao;
                    ementaInserir.Republicacao = string.Empty;
                    ementaInserir.Ementa = titulo;
                    ementaInserir.TituloAto = titulo;
                    ementaInserir.Especie = especie;
                    ementaInserir.NumeroAto = numero;
                    ementaInserir.DataEdicao = string.Empty;
                    ementaInserir.Texto = Regex.Replace(item.Trim(), @"\s+", " ");
                    ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(item)));

                    ementaInserir.Escopo = "FED";
                    ementaInserir.IdFila = listaUrl[0].Lista_Nivel2[0].Id;

                    itemListaVez.ListaEmenta.Add(ementaInserir);

                    dynamic itemFonte = new ExpandoObject();

                    itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                });
            }

            #endregion

            #endregion

            #region "Sefaz Rio Grande do Sul - Lote 4"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                webBrowser1 = new WebBrowser();

                webBrowser1.ScriptErrorsSuppressed = true;

                webBrowser1.Navigate("http://www.legislacao.sefaz.rs.gov.br/Site/Search.aspx?a=&CodArea=0");

                webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted_SefazRs);
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazRS"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazRS");

                string urlTratada = string.Empty;

                int inicioCaptura = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (inicioCaptura < 3675)
                            {
                                inicioCaptura++;
                                continue;
                            }

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim();

                            itemListaVez.Url = urlTratada;

                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            /*Fase 1 obter a url para o documento final*/

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, string.Empty);

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            respostaNivel2_ = respostaNivel2_.Substring(respostaNivel2_.IndexOf("id=FrameDoc"));

                            respostaNivel2_ = respostaNivel2_.Substring(respostaNivel2_.IndexOf("src=") + 4, respostaNivel2_.IndexOf(">") - (respostaNivel2_.IndexOf("src=") + 4));

                            urlTratada = urlTratada.Substring(0, urlTratada.LastIndexOf("/") + 1) + respostaNivel2_;

                            /*Fase 2 obter o documento final*/

                            respostaNivel2_ = getHtmlPaginaByGet(urlTratada, "default");

                            /*Captura css*/

                            var listaCss = Regex.Split(respostaNivel2_.Replace("\"", string.Empty).Replace("\n", string.Empty), "<link").ToList();

                            listaCss.RemoveAt(0);

                            string cssList = string.Empty;

                            if (listaCss.Count > 0)
                            {
                                listaCssTratada = new List<string>();

                                listaCss.ForEach(delegate(string x)
                                {
                                    string novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://www.legislacao.sefaz.rs.gov.br" + (novaCss.Contains("..") ? string.Empty : "/"))).Replace("..", string.Empty);

                                    listaCssTratada.Add(novaCss);
                                });

                                listaCssTratada.ForEach(x => cssList += x);
                            }

                            /*Fim Captura css*/

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            string textFull = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<div class=struct titulo_container"), respostaNivel2_.IndexOf("<div id=DivFooter") - respostaNivel2_.IndexOf("<div class=struct titulo_container"));

                            int corte = textFull.IndexOf("<br>") >= 0 && textFull.IndexOf("<br>") < textFull.IndexOf("</p>") ? textFull.IndexOf("<br>") : textFull.IndexOf("</p>");

                            string titulo = ObterStringLimpa(textFull.Substring(0, corte));

                            string numero = titulo.ToLower().Trim().Contains("/") ? titulo.Trim().Substring(0, titulo.ToLower().Trim().IndexOf("/")) :
                                                titulo.ToLower().Trim().Contains(",") ? titulo.Trim().Substring(0, titulo.ToLower().Trim().IndexOf(",")) :
                                                    titulo.ToLower().Trim().Contains(" de") ? titulo.Trim().Substring(0, titulo.ToLower().Trim().IndexOf(" de")) : titulo.Trim();

                            //string especie = titulo.Substring(0, titulo.ToUpper().Replace("NOVEMBRO", string.Empty).LastIndexOf(" N"));
                            int indexEspecie = titulo.ToLower().Replace("novembro", string.Empty).LastIndexOf(" n") > 0 ? titulo.ToLower().Replace("novembro", string.Empty).LastIndexOf(" n") : obterPontoCorte(titulo);
                            string especie = indexEspecie >= 0 ? titulo.Substring(0, indexEspecie) : string.Empty;

                            string ementa = string.Empty;

                            if (textFull.IndexOf("<div class=struct ementa_container") > 0)
                                ementa = textFull.Substring(textFull.IndexOf("<div class=struct ementa_container")).Substring(0, textFull.Substring(textFull.IndexOf("<div class=struct ementa_container")).IndexOf("</div>"));

                            string publicacao = string.Empty;

                            if (textFull.IndexOf("class=reference>") > 0)
                            {
                                publicacao = textFull.Substring(textFull.IndexOf("class=reference>"));
                                publicacao = ObterStringLimpa("<" + publicacao.Substring(0, publicacao.IndexOf("</span>")));
                            }

                            textFull = cssList + textFull;

                            /**Tratamento para inserir na base**/
                            dynamic ementaInserir = new ExpandoObject();

                            /**Outros**/
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /**Arquivo**/
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "RS") + "}";

                            /**Default**/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ObterStringLimpa(ementa);
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = textFull;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(textFull)));

                            ementaInserir.Escopo = "RS";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            ex.Source = "Docs > Sefaz RS > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sef Santa Catarina - Lote 4"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                List<string> listaUrls = new List<string>() {"http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/leis/indice_leis_complementares.htm#leis_complementar",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/leis/indice_leis.htm#leis_2009",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Medidas_Provisorias/indice_mp.htm",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Decretos/indice_Decretos.htm",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Portarias/indice_portarias.htm",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Atos_Diat/indice_atos_diat.htm",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Atos_Diat/indice_Atos_Homologatorios.htm#Atos_Homologatorios_2008",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Atos_Diat/indice_Atos_Homolog_MVC.htm#Atos_Homolog_MVC_2016",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Consultas/Ind_RN.htm",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Consultas/Ind_Consultas_Analitico.htm",
                "http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Notas_Tecnicas/indice_Notas_Tecnicas.htm#ntec_2015﻿"};

                listaInserir = new List<dynamic>();

                string resposta = string.Empty;

                foreach (var url in listaUrls)
                {
                    Thread.Sleep(2000);

                    resposta = string.Empty;

                    if (url.Equals("http://legislacao.sef.sc.gov.br/legtrib_internet/Indices/Consultas/Ind_RN.htm"))
                        resposta = getHtmlPaginaByGet(url, "default");
                    else
                        resposta = getHtmlPaginaByGet(url, string.Empty);

                    objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Secretaria do estado da fazenda de SC";
                    objUrll.Url = url;
                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("\"", " "));

                    var listaUrlDocs = Regex.Split(resposta, "href=").ToList();

                    listaUrlDocs.RemoveAt(0);

                    listaUrlDocs.ForEach(delegate(string x)
                    {
                        var urlTratada = x.Replace("../../", "http://legislacao.sef.sc.gov.br/").Substring(0, x.Replace("../../", "http://legislacao.sef.sc.gov.br/").IndexOf(">"));

                        objItenUrl = new ExpandoObject();

                        objItenUrl.Url = urlTratada;

                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    });

                    listaInserir.Add(objUrll);
                }

                new BuscaLegalDao().AtualizarFontes(listaInserir);
            }

            #endregion

            #region "Captura DOC's"

            #region "Modelo Doc 1"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefSC"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefSC");

                string urlTratada = string.Empty;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim();

                            if (urlTratada.ToLower().Contains("planalto"))
                                continue;

                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            /*Fase 1 obter a url para o documento final*/

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, "default");

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            /*Captura css*/

                            var listaCss = Regex.Split(respostaNivel2_.Replace("\"", string.Empty).Replace("\n", string.Empty), "<link").ToList();

                            listaCss.RemoveAt(0);

                            string cssList = string.Empty;

                            if (listaCss.Count > 0)
                            {
                                listaCssTratada = new List<string>();

                                listaCss.ForEach(delegate(string x)
                                {
                                    string novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("Incluir Raiz do Site" + (novaCss.Contains("..") ? string.Empty : "/"))).Replace("..", string.Empty);

                                    listaCssTratada.Add(novaCss);
                                });

                                listaCssTratada.ForEach(x => cssList += x);
                            }

                            /*Fim Captura css*/

                            string descStyle = string.Empty;

                            if (respostaNivel2_.Contains("<style"))
                                descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                            var texto = respostaNivel2_.Substring(respostaNivel2_.ToLower().IndexOf("<body"))
                                        .Substring(0, respostaNivel2_.Substring(respostaNivel2_.ToLower().IndexOf("<body")).ToLower().IndexOf("</body>") + 7);

                            int indexTitulo = respostaNivel2_.ToLower().IndexOf("class=document") > 0 ? respostaNivel2_.ToLower().IndexOf("class=document") : respostaNivel2_.ToLower().IndexOf("class=01documento");

                            if (indexTitulo < 0)
                                indexTitulo = respostaNivel2_.ToLower().IndexOf("<p class=msonormal align=center") >= 0 &&
                                                (respostaNivel2_.ToLower().IndexOf("<p class=msonormal align=center") < respostaNivel2_.ToLower().IndexOf("<p class=msobodytext2 align=center") ||
                                                    respostaNivel2_.ToLower().IndexOf("<p class=msobodytext2 align=center") < 0) ?
                                                        respostaNivel2_.ToLower().IndexOf("<p class=msonormal align=center") : respostaNivel2_.ToLower().IndexOf("<p class=msobodytext2 align=center") > 0 ?
                                                            respostaNivel2_.ToLower().IndexOf("<p class=msobodytext2 align=center") : respostaNivel2_.ToLower().IndexOf("<p class=msonormal>");

                            string titulo = string.Empty;
                            string ementa = string.Empty;

                            if (indexTitulo > 0)
                                titulo = "<" + respostaNivel2_.Substring(indexTitulo).Substring(0, respostaNivel2_.Substring(indexTitulo).IndexOf("</p>"));

                            else
                            {
                                indexTitulo = respostaNivel2_.ToLower().IndexOf("<span style=mso-bookmark:ole_link");
                                titulo = "<" + respostaNivel2_.Substring(indexTitulo).Substring(0, respostaNivel2_.Substring(indexTitulo).IndexOf("</span>"));

                                var list1 = Regex.Split(respostaNivel2_.ToLower(), "<p class=estilo style=").ToList();

                                list1.RemoveAt(0);
                                list1.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                list1.ForEach(x => ementa = ementa.Equals(string.Empty) ? ObterStringLimpa("<" + x.Substring(0, x.IndexOf("</p>"))) : ementa);
                            }

                            titulo = ObterStringLimpa(titulo);

                            if (titulo.Trim().Equals(string.Empty) && respostaNivel2_.ToLower().Contains("class=document"))
                            {
                                var list1 = Regex.Split(respostaNivel2_.ToLower(), "class=document").ToList();

                                list1.RemoveAt(0);
                                list1.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                list1.ForEach(x => titulo = titulo.Equals(string.Empty) ? ObterStringLimpa("<" + x.Substring(0, x.IndexOf("</p>"))) : titulo);
                            }

                            if (titulo.Trim().Equals(string.Empty) && respostaNivel2_.ToLower().Contains("class=01documento"))
                            {
                                var list1 = Regex.Split(respostaNivel2_.ToLower(), "class=01documento").ToList();

                                list1.RemoveAt(0);
                                list1.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                list1.ForEach(x => titulo = titulo.Equals(string.Empty) ? ObterStringLimpa("<" + x.Substring(0, x.IndexOf("</p>"))) : titulo);
                            }

                            int indexPublicacao = respostaNivel2_.ToLower().IndexOf("<p class=publicado") >= 0 ? respostaNivel2_.ToLower().IndexOf("<p class=publicado") : respostaNivel2_.ToLower().IndexOf("<p class=msonormal");
                            string publicacao = indexPublicacao > 0 ? respostaNivel2_.Substring(indexPublicacao).Substring(0, respostaNivel2_.ToLower().Substring(indexPublicacao).IndexOf("</p>")) : string.Empty;

                            int indexCorte = respostaNivel2_.ToLower().IndexOf("<p class=ementa") >= 0 ? respostaNivel2_.ToLower().IndexOf("<p class=ementa") : respostaNivel2_.ToLower().IndexOf("<p class=msobodytextindent2>");
                            ementa = indexCorte >= 0 && ementa.Equals(string.Empty) ? respostaNivel2_.Substring(indexCorte).Substring(0, respostaNivel2_.Substring(indexCorte).IndexOf("</p>")) : ementa;

                            if (ementa.Equals(string.Empty) &&
                                   respostaNivel2_.ToLower().IndexOf("class=document") < 0 &&
                                        respostaNivel2_.ToLower().IndexOf("class=01documento") < 0 &&
                                            respostaNivel2_.ToLower().IndexOf("<p class=msobodytext2 align=center") < 0 &&
                                                respostaNivel2_.ToLower().IndexOf("<p class=msonormal align=center") < 0 &&
                                                    respostaNivel2_.ToLower().IndexOf("<p class=msonormal>") > 0)
                            {
                                var list1 = Regex.Split(respostaNivel2_.ToLower(), "<p class=msonormal>").ToList();

                                list1.RemoveRange(0, 2);
                                list1.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                list1.ForEach(x => ementa = ementa.Equals(string.Empty) ? ObterStringLimpa("<" + x.Substring(0, x.IndexOf("</p>"))) : ementa);
                            }

                            string numero = titulo.IndexOf("/") >= 0 && (titulo.IndexOf("/") < titulo.IndexOf(",") || titulo.IndexOf(",") < 0) ?
                                                titulo.Substring(0, titulo.LastIndexOf("/")) : titulo.IndexOf(",") > 0 ?
                                                    titulo.Substring(0, titulo.IndexOf(",")) : titulo.ToLower().IndexOf(" de") > 0 ?
                                                        titulo.Substring(0, titulo.ToLower().IndexOf(" de")) : titulo;

                            int indexEspecie = titulo.ToLower().Replace("novembro", string.Empty).LastIndexOf(" n") > 0 ? titulo.ToLower().Replace("novembro", string.Empty).LastIndexOf(" n") : obterPontoCorte(titulo);
                            string especie = indexEspecie >= 0 ? titulo.Substring(0, indexEspecie) : string.Empty;

                            publicacao = ObterStringLimpa(publicacao);
                            ementa = ObterStringLimpa(ementa);
                            numero = ObterStringLimpa(numero);
                            especie = ObterStringLimpa(especie);

                            /**Tratamento para inserir na base**/
                            dynamic ementaInserir = new ExpandoObject();

                            /**Outros**/
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /**Arquivo**/
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "SC") + "}";

                            /**Default**/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = descStyle + texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(texto)));

                            ementaInserir.Escopo = "SC";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();
                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };
                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            ex.Source = "Docs > Sefaz SC > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #region "Modelo Doc 2"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefSC2"))
            {
                listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefSC2");

                string urlTratada = string.Empty;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim();

                            if (urlTratada.ToLower().Contains("planalto"))
                                continue;

                            itemListaVez.Url = urlTratada;

                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            /*Fase 1 obter a url para o documento final*/

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, "default");

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            /*Captura css*/

                            var listaCss = Regex.Split(respostaNivel2_.Replace("\"", string.Empty).Replace("\n", string.Empty), "<link").ToList();

                            listaCss.RemoveAt(0);

                            string cssList = string.Empty;

                            if (listaCss.Count > 0)
                            {
                                listaCssTratada = new List<string>();

                                listaCss.ForEach(delegate(string x)
                                {
                                    string novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("Incluir Raiz do Site" + (novaCss.Contains("..") ? string.Empty : "/"))).Replace("..", string.Empty);

                                    listaCssTratada.Add(novaCss);
                                });

                                listaCssTratada.ForEach(x => cssList += x);
                            }

                            /*Fim Captura css*/

                            string descStyle = string.Empty;

                            if (respostaNivel2_.Contains("<style"))
                                descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                            var texto = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<body"))
                                        .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<body")).IndexOf("</body>") + 7);

                            string titulo = string.Empty;

                            if (texto.ToLower().IndexOf("class=consulta") < 0 && texto.ToLower().IndexOf("class=redaoatual ") < 0 && texto.ToLower().IndexOf("class=msoheader") < 0 && texto.ToLower().IndexOf("class=resoluoassunto") < 0 && texto.ToLower().IndexOf("class=publicado") < 0)
                                titulo = obterDadosColection(texto, "consulta nº:");

                            else if (texto.ToLower().IndexOf("class=consulta") > 0)
                                titulo = "<" + texto.ToLower().Substring(texto.ToLower().IndexOf("class=consulta")).Substring(0, texto.ToLower().Substring(texto.ToLower().IndexOf("class=consulta")).IndexOf("</p>"));

                            else if (texto.ToLower().IndexOf("class=msonormal") > 0 && (texto.ToLower().IndexOf("class=msonormal") < texto.ToLower().IndexOf("class=msoheader") || texto.ToLower().IndexOf("class=msoheader") < 0) && (texto.ToLower().IndexOf("class=msonormal") < texto.ToLower().IndexOf("class=redaoatual ") || texto.ToLower().IndexOf("class=redaoatual ") < 0) && (texto.ToLower().IndexOf("class=msonormal") < texto.ToLower().IndexOf("class=resoluoassunto") || texto.ToLower().IndexOf("class=resoluoassunto") < 0))
                            {
                                if (texto.ToLower().IndexOf("class=msonormal") < texto.ToLower().IndexOf("class=publicado") && texto.ToLower().IndexOf("class=publicado") > 0 && texto.ToLower().IndexOf("class=resoluoassunto") < 0)
                                    titulo = texto.ToLower().Substring(texto.ToLower().IndexOf("class=msonormal")).Substring(texto.ToLower().Substring(texto.ToLower().IndexOf("class=msonormal")).IndexOf("</div>"), texto.ToLower().Substring(texto.ToLower().IndexOf("class=msonormal")).IndexOf("class=publicado") - texto.Substring(texto.ToLower().IndexOf("class=msonormal")).IndexOf("</div>"));

                                else
                                {
                                    titulo = "<" + texto.ToLower().Substring(texto.ToLower().IndexOf("class=msonormal")).Substring(0, texto.ToLower().Substring(texto.ToLower().IndexOf("class=msonormal")).IndexOf("</p>"));

                                    if (string.IsNullOrEmpty(ObterStringLimpa(titulo.Trim())))
                                    {
                                        var textoAux = texto.Remove(texto.ToLower().IndexOf("class=msonormal"), (texto.ToLower().IndexOf("class=msonormal") + "class=msonormal".Length - texto.ToLower().IndexOf("class=msonormal")));

                                        if (textoAux.ToLower().IndexOf("class=msonormal") > textoAux.ToLower().IndexOf("class=msoheader") && textoAux.ToLower().IndexOf("class=msoheader") >= 0)
                                            titulo = "<" + textoAux.ToLower().Substring(textoAux.ToLower().IndexOf("class=msoheader")).Substring(0, textoAux.ToLower().Substring(textoAux.ToLower().IndexOf("class=msoheader")).IndexOf("</p>"));
                                        else
                                            titulo = "<" + textoAux.ToLower().Substring(textoAux.ToLower().IndexOf("class=msonormal")).Substring(0, textoAux.ToLower().Substring(textoAux.ToLower().IndexOf("class=msonormal")).IndexOf("</p>"));
                                    }
                                }
                            }
                            else if (texto.ToLower().IndexOf("class=redaoatual ") > 0 && (texto.ToLower().IndexOf("class=redaoatual ") < texto.ToLower().IndexOf("class=msoheader") || texto.ToLower().IndexOf("class=msoheader") < 0))
                                titulo = "<" + texto.ToLower().Substring(texto.ToLower().IndexOf("class=redaoatual ")).Substring(0, texto.ToLower().Substring(texto.ToLower().IndexOf("class=redaoatual ")).IndexOf("</p>"));

                            else if (texto.ToLower().IndexOf("class=msoheader") > 0)
                                titulo = "<" + texto.ToLower().Substring(texto.ToLower().IndexOf("class=msoheader")).Substring(0, texto.ToLower().Substring(texto.ToLower().IndexOf("class=msoheader")).IndexOf("</p>"));

                            else if (texto.ToLower().IndexOf("class=resoluoassunto") > 0 && texto.ToLower().IndexOf("class=publicado") > 0)
                                titulo = texto.ToLower().Substring(texto.ToLower().IndexOf("class=resoluoassunto")).Substring(texto.ToLower().Substring(texto.ToLower().IndexOf("class=resoluoassunto")).IndexOf("</div>"), texto.ToLower().Substring(texto.ToLower().IndexOf("class=resoluoassunto")).IndexOf("class=publicado") - texto.Substring(texto.ToLower().IndexOf("class=resoluoassunto")).IndexOf("</div>"));

                            titulo = ObterStringLimpa(titulo);

                            string especie = string.Empty;
                            string numero = string.Empty;

                            if (string.IsNullOrEmpty(titulo))
                                numero = "0";

                            else
                            {
                                especie = titulo.ToLower().IndexOf(" n") > 0 ? titulo.Substring(0, titulo.ToLower().IndexOf(" n")) : titulo.IndexOf(":") > 0 ? titulo.Substring(0, titulo.IndexOf(":")) : titulo.Substring(0, obterPontoCorte(titulo));
                                numero = titulo.IndexOf("/") >= 0 && (titulo.IndexOf("/") < titulo.IndexOf(",") || titulo.IndexOf(",") < 0) ? titulo.Substring(0, titulo.IndexOf("/")) : titulo.Substring(0, titulo.IndexOf(","));
                            }

                            string publicacao = string.Empty;
                            string ementa = string.Empty;

                            if (texto.ToLower().IndexOf("class=consulta") < 0 && texto.ToLower().IndexOf("class=redaoatual ") < 0 && texto.ToLower().IndexOf("class=msoheader") < 0 && texto.ToLower().IndexOf("class=resoluoassunto") < 0 && texto.ToLower().IndexOf("class=publicado") < 0)
                            {
                                publicacao = obterDadosColection(texto, "d.o.e.");
                                ementa = obterDadosColection(texto, "ementa:");
                            }
                            else
                            {
                                publicacao = texto.ToLower().IndexOf("class=publicado") > 0 ? "<" + texto.Substring(texto.ToLower().IndexOf("class=publicado")).Substring(0, texto.Substring(texto.ToLower().IndexOf("class=publicado")).IndexOf("</p>")) : string.Empty;
                                ementa = texto.ToLower().IndexOf("class=resoluoassunto") > 0 ? obterEmentaTexto(texto) : string.Empty;
                            }

                            especie = ObterStringLimpa(especie);
                            numero = Regex.Replace(ObterStringLimpa(numero), "[^0-9]+", string.Empty);
                            publicacao = ObterStringLimpa(publicacao);
                            ementa = ObterStringLimpa(ementa);

                            texto = removeTagScript(texto.Trim());

                            /**Tratamento para inserir na base**/
                            dynamic ementaInserir = new ExpandoObject();

                            /**Outros**/
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /**Arquivo**/
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "SC") + "}";

                            /**Default**/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = numero;
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = descStyle + Regex.Replace(texto.Trim(), @"\s+", " ");
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(texto)));

                            ementaInserir.Escopo = "SC";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }

            #endregion

            #region "Modelo Doc 3"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefSC3"))
            {
                listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefSC3");

                string urlTratada = string.Empty;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim();

                            if (urlTratada.ToLower().Contains("planalto"))
                                continue;

                            itemListaVez.Url = urlTratada;

                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            /*Fase 1 obter a url para o documento final*/

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, "default");

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            /*Captura css*/

                            var listaCss = Regex.Split(respostaNivel2_.Replace("\"", string.Empty).Replace("\n", string.Empty), "<link").ToList();

                            listaCss.RemoveAt(0);

                            string cssList = string.Empty;

                            if (listaCss.Count > 0)
                            {
                                listaCssTratada = new List<string>();

                                listaCss.ForEach(delegate(string x)
                                {
                                    string novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("Incluir Raiz do Site" + (novaCss.Contains("..") ? string.Empty : "/"))).Replace("..", string.Empty);

                                    listaCssTratada.Add(novaCss);
                                });

                                listaCssTratada.ForEach(x => cssList += x);
                            }

                            /*Fim Captura css*/

                            string descStyle = string.Empty;

                            if (respostaNivel2_.Contains("<style"))
                                descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                            var texto = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<body"))
                                            .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<body")).IndexOf("</body>"));

                            string titulo = string.Empty;
                            string ementa = string.Empty;

                            if (texto.ToLower().IndexOf("class=documento01") > 0)
                            {
                                titulo = "<" + texto.Substring(texto.ToLower().IndexOf("class=documento01")).Substring(0, texto.Substring(texto.ToLower().IndexOf("class=documento01")).IndexOf("</p>"));

                                if (ObterStringLimpa(titulo).Equals(string.Empty))
                                    titulo = "<" + texto.Substring(texto.ToLower().LastIndexOf("class=documento01")).Substring(0, texto.Substring(texto.ToLower().LastIndexOf("class=documento01")).IndexOf("</p>"));

                                texto = "<" + texto.Substring(texto.ToLower().IndexOf("class=documento01"));

                                if (texto.ToLower().Contains("<p class=ementa"))
                                    ementa = texto.Substring(texto.ToLower().IndexOf("<p class=ementa")).Substring(0, texto.Substring(texto.ToLower().IndexOf("<p class=ementa")).IndexOf("</p>"));

                                else
                                {
                                    var listaEmenta = Regex.Split(texto, "<p class=MsoNormal align=center").ToList();

                                    listaEmenta.RemoveAt(0);

                                    foreach (string item in listaEmenta)
                                    {
                                        if (!ObterStringLimpa("<" + item).Trim().Equals(string.Empty))
                                        {
                                            ementa = "<" + item.Substring(0, item.IndexOf("</p>"));
                                            break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                titulo = texto.Substring(texto.IndexOf("<p class=MsoNormal align=center")).Substring(0, texto.Substring(texto.IndexOf("<p class=MsoNormal align=center")).IndexOf("</p>"));
                                texto = texto.Substring(texto.IndexOf("<p class=MsoNormal align=center"));

                                ementa = texto.IndexOf("<b style=mso-bidi-font-weight:") > 0 ?
                                                texto.Substring(texto.IndexOf("<b style=mso-bidi-font-weight:")).Substring(0, texto.Substring(texto.IndexOf("<b style=mso-bidi-font-weight:")).IndexOf("</b>")) :
                                                    texto.Substring(texto.IndexOf("<b style='mso-bidi-font-weight:")).Substring(0, texto.Substring(texto.IndexOf("<b style='mso-bidi-font-weight:")).IndexOf("</b>"));
                            }

                            string numero = string.Empty;
                            string especie = string.Empty;

                            titulo = ObterStringLimpa(titulo).Trim();

                            if (string.IsNullOrEmpty(titulo))
                                numero = "0";

                            else
                            {
                                especie = titulo.ToLower().IndexOf(" n") > 0 ? titulo.Substring(0, titulo.ToLower().IndexOf(" n")) : titulo.IndexOf(":") > 0 ? titulo.Substring(0, titulo.IndexOf(":")) : titulo.Substring(0, obterPontoCorte(titulo));
                                numero = titulo.IndexOf("/") >= 0 && (titulo.IndexOf("/") < titulo.IndexOf(",") || titulo.IndexOf(",") < 0) ? titulo.Substring(0, titulo.IndexOf("/")) : titulo.IndexOf(",") >= 0 ? titulo.Substring(0, titulo.IndexOf(",")) : titulo;
                            }

                            especie = ObterStringLimpa(especie);
                            numero = Regex.Replace(ObterStringLimpa(numero), "[^0-9]+", string.Empty);
                            ementa = ObterStringLimpa(ementa);

                            texto = removeTagScript(texto.Trim());

                            /**Tratamento para inserir na base**/
                            dynamic ementaInserir = new ExpandoObject();

                            /**Outros**/
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /**Arquivo**/
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "SC") + "}";

                            /**Default**/
                            ementaInserir.Publicacao = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = numero;
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = descStyle + Regex.Replace(texto.Trim(), @"\s+", " ");
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(texto)));

                            ementaInserir.Escopo = "SC";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });

                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }

            #endregion

            #endregion

            #endregion

            #region "Sefaz PR"

            #region "URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                objUrll = new ExpandoObject();

                objUrll.Indexacao = "Secretaria do estado da fazenda do PR";
                objUrll.Url = "http://www.fazenda.pr.gov.br/modules/conteudo/conteudo.php?conteudo=248";
                objUrll.Lista_Nivel2 = new List<dynamic>();

                objItenUrl = new ExpandoObject();

                objItenUrl.Url = "https://www.arinternet.pr.gov.br/portalsefa/_l_DownloadLegislacao2.asp?eTpDoc=0&eTpPer=9&eDtPublicacaoIni=&eDtPublicacaoFim=&eNrDocumento=&eAnoDocumento=&eTpMod=1";
                objUrll.Lista_Nivel2.Add(objItenUrl);

                new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
            }

            #endregion

            #region "DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazPR"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazPR");

                string urlTratada = string.Empty;

                int ContinueLn = 0;

                dynamic ementaInserir;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            urlTratada = itemLista_Nivel2.Url.Trim();

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, "default");

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            respostaNivel2_ = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<table  border=0 width=670 cellpadding=2 height=30 BGCOLOR=")).Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<table  border=0 width=670 cellpadding=2 height=30 BGCOLOR=")).IndexOf("</table>"));

                            var listaTR = Regex.Split(respostaNivel2_, "</tr>").ToList();

                            listaTR.RemoveRange(0, 2);
                            listaTR.RemoveAt(listaTR.Count - 1);

                            /*Quebra e fazer o Loop dos Links do PDF*/

                            foreach (string item in listaTR)
                            {
                                try
                                {
                                    if (ContinueLn < 1475)
                                    {
                                        ContinueLn++;
                                        continue;
                                    }

                                    dynamic itemListaVez = new ExpandoObject();

                                    itemListaVez.ListaEmenta = new List<dynamic>();
                                    urlTratada = itemLista_Nivel2.Url.Trim();
                                    itemListaVez.Url = urlTratada;
                                    itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                                    ContinueLn++;

                                    /*Composição do Item*/
                                    var listaItens = Regex.Split(item, "</td>").ToList();

                                    string UrlArquivo = listaItens[1].Substring(listaItens[1].ToLower().IndexOf("href=") + 5).Substring(0, listaItens[1].Substring(listaItens[1].ToLower().IndexOf("href=") + 5).IndexOf(" "));

                                    /** Arquivo **/
                                    string nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), UrlArquivo.Substring(UrlArquivo.LastIndexOf("/") + 1));

                                    using (WebClient webClient = new WebClient())
                                    {
                                        webClient.DownloadFile(UrlArquivo, nomeArq);
                                    }

                                    byte[] arrayFile = File.ReadAllBytes(nomeArq);
                                    string conteudoPdf = LeArquivo(nomeArq);
                                    File.Delete(nomeArq);

                                    /** Fim-Arquivo **/
                                    string hash = GerarHash(removerCaracterEspecial(Regex.Replace(conteudoPdf.Replace("\n", string.Empty).Replace("\t", string.Empty).Trim(), @"\s+", " ")));

                                    ementaInserir = new ExpandoObject();

                                    /** Outros **/
                                    ementaInserir.DescSigla = string.Empty;
                                    ementaInserir.HasContent = false;

                                    /** Arquivo **/
                                    ementaInserir.ListaArquivos = new List<ArquivoUpload>();
                                    ementaInserir.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = nomeArq.Substring(nomeArq.LastIndexOf(".") + 1), NomeArquivo = urlTratada.Split('|')[0].Substring(urlTratada.Split('|')[0].LastIndexOf("/") + 1) });

                                    ementaInserir.Tipo = 3;
                                    ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "PR") + "}";

                                    string numero = ObterStringLimpa(listaItens[1]).Contains("/") ? ObterStringLimpa(listaItens[1]).Substring(0, ObterStringLimpa(listaItens[1]).IndexOf("/")) : ObterStringLimpa(listaItens[1]);

                                    /** Default **/
                                    ementaInserir.Publicacao = ObterStringLimpa(listaItens[2]);
                                    ementaInserir.Sigla = string.Empty;
                                    ementaInserir.Republicacao = string.Empty;
                                    ementaInserir.Ementa = ObterStringLimpa(listaItens[4]);
                                    ementaInserir.TituloAto = string.Format("{0} {1} {2}", ObterStringLimpa(listaItens[0]), ObterStringLimpa(listaItens[1]), ObterStringLimpa(listaItens[3]));
                                    ementaInserir.Especie = ObterStringLimpa(listaItens[0]);
                                    ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                                    ementaInserir.DataEdicao = string.Empty;
                                    ementaInserir.Texto = Regex.Replace(conteudoPdf.Replace("\0", string.Empty), @"\s+", " ");
                                    ementaInserir.Hash = hash;

                                    ementaInserir.Escopo = "PR";
                                    ementaInserir.IdFila = itemLista_Nivel2.Id;

                                    itemListaVez.ListaEmenta.Add(ementaInserir);

                                    dynamic itemFonte = new ExpandoObject();

                                    itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                                }
                                catch (Exception ex)
                                {
                                    ex.Source = "Docs > sefazPR > lv1";
                                    new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            ex.Source = "Docs > sefazPR > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Confaz"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                List<string> listUrlDocs = new List<string>() {/*"https://www.confaz.fazenda.gov.br/legislacao/convenios-ecf"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/convenios"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/convenios_arrecadacao"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/copy_of_convenios-diversos"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/ajustes"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/atos"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/atos-declaratorios"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/atos-pmpf"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/atos-mva"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/atos_confaz"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/despacho"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/portarias-1"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/protocolos"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/protocolos-ecf"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/protocolos-ipva"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/protocolos-diversos"
                                                               ,*/"https://www.confaz.fazenda.gov.br/legislacao/aliquotas-icms-estaduais"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/regimento"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/orgaos-credenciados"
                                                               ,"https://www.confaz.fazenda.gov.br/legislacao/arquivo-manuais"};

                foreach (var itemUrl in listUrlDocs)
                {
                    objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Conselho Nacional de Política Fazendária";
                    objUrll.Url = itemUrl;
                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    if (itemUrl.Contains("/aliquotas-icms-estaduais") || itemUrl.Contains("/regimento") || itemUrl.Contains("/orgaos-credenciados") || itemUrl.Contains("/arquivo-manuais"))
                    {
                        string resposta = getHtmlPaginaByGet(objUrll.Url, string.Empty);

                        resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("\"", string.Empty));

                        resposta = resposta.Substring(resposta.IndexOf("<div id=content-core")).Substring(0, resposta.Substring(resposta.IndexOf("<div id=content-core")).IndexOf("<div id=viewlet-below-content"));

                        var listaUrlDocs = Regex.Split(resposta, "href=").ToList().Select(x => x.Substring(0, (x.IndexOf(" ") < x.IndexOf(">") ? x.IndexOf(" ") : x.IndexOf(">")))).ToList();

                        listaUrlDocs.RemoveAt(0);
                        listaUrlDocs.RemoveAll(x => x.Contains("../") || x.Contains("/convenio-icms/") || x.Contains("/ajustesinief_"));

                        listaUrlDocs.ForEach(delegate(string item)
                        {
                            objItenUrl = new ExpandoObject();

                            objItenUrl.Url = item;
                            objUrll.Lista_Nivel2.Add(objItenUrl);
                        });
                    }
                    else
                    {
                        string resposta = getHtmlPaginaByGet(objUrll.Url, string.Empty);

                        resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("\"", " "));

                        resposta = resposta.Substring(resposta.IndexOf("<div id=parent-fieldname-text>")).Substring(0, resposta.Substring(resposta.IndexOf("<div id=parent-fieldname-text>")).IndexOf("<div id=viewlet-below-content>"));

                        var listaUrlDocs = Regex.Split(resposta, "href=").ToList().Select(x => x.Substring(0, (x.IndexOf(" ") < x.IndexOf(">") ? x.IndexOf(" ") : x.IndexOf(">")))).ToList();

                        listaUrlDocs.RemoveAt(0);
                        listaUrlDocs.RemoveAll(x => x.Contains("../") || x.Contains("/convenio-icms/") || x.Contains("/ajustesinief_"));

                        foreach (var item in listaUrlDocs)
                        {
                            resposta = getHtmlPaginaByGet(item, string.Empty);

                            resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", string.Empty).Replace("\r", string.Empty).Replace("\t", string.Empty).Replace("\"", string.Empty));

                            int indexCorte = resposta.IndexOf("<div id=viewlet-below-content>");

                            if (indexCorte > 0)
                                resposta = resposta.Substring(resposta.IndexOf("<div id=parent-fieldname-text>")).Substring(0, resposta.Substring(resposta.IndexOf("<div id=parent-fieldname-text>")).IndexOf("<div id=viewlet-below-content>"));
                            else
                                resposta = resposta.Substring(resposta.IndexOf("<div id=parent-fieldname-text>")).Substring(0, resposta.Substring(resposta.IndexOf("<div id=parent-fieldname-text>")).IndexOf("</table>"));

                            if (resposta.ToLower().Contains("<table class=plain") || resposta.ToLower().Contains("<table class=msonormaltable>") || resposta.ToLower().Contains("<table class=msonormaltable"))
                            {
                                var listaUrlDocsN2 = Regex.Split(resposta, "href=").ToList().Select(x => x.Substring(0, (x.IndexOf(" ") < x.IndexOf(">") ? x.IndexOf(" ") : x.IndexOf(">")))).ToList();

                                listaUrlDocsN2.RemoveAt(0);
                                listaUrlDocsN2.RemoveAll(x => x.Contains("../"));

                                listaUrlDocsN2.ForEach(delegate(string itemN2)
                                {
                                    objItenUrl = new ExpandoObject();

                                    objItenUrl.Url = itemN2.Contains("https://www.confaz.fazenda.gov.br") ? itemN2 : "https://www.confaz.fazenda.gov.br" + itemN2;
                                    objUrll.Lista_Nivel2.Add(objItenUrl);
                                });
                            }
                            else
                            {
                                objItenUrl = new ExpandoObject();

                                objItenUrl.Url = item;
                                objUrll.Lista_Nivel2.Add(objItenUrl);
                            }
                        }
                    }

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("confaz"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("confaz");

                string urlTratada = string.Empty;

                dynamic ementaInserir;
                int contX = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (contX < 0)
                            {
                                contX++;
                                continue;
                            }

                            contX++;

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, string.Empty);

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", string.Empty).Replace("\r", string.Empty).Replace("\t", string.Empty));

                            string hash = string.Empty;
                            string texto = string.Empty;

                            ementaInserir = new ExpandoObject();
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string ementa = string.Empty;

                            if (urlTratada.ToLower().Contains(".pdf") || urlTratada.ToLower().Contains(".xlsx"))
                            {
                                /** Arquivo **/
                                string nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), "Extrair o nome do arquivo");

                                using (WebClient webClient = new WebClient())
                                {
                                    webClient.DownloadFile("Incluir Url do Arquivo", nomeArq);
                                }

                                byte[] arrayFile = File.ReadAllBytes(nomeArq);
                                string conteudoPdf = LeArquivo(nomeArq);
                                File.Delete(nomeArq);

                                ementaInserir.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = nomeArq.Substring(nomeArq.LastIndexOf(".") + 1), NomeArquivo = urlTratada.Split('|')[0].Substring(urlTratada.Split('|')[0].LastIndexOf("/") + 1) });
                                /** Fim-Arquivo **/

                                hash = GerarHash(removerCaracterEspecial(Regex.Replace(conteudoPdf.Replace("\n", string.Empty).Replace("\t", string.Empty).Trim(), @"\s+", " ")));
                            }
                            else
                            {
                                texto = respostaNivel2_.ToLower().Substring(respostaNivel2_.ToLower().IndexOf("<div id=content-core>")).Substring(0, respostaNivel2_.ToLower().Substring(respostaNivel2_.ToLower().IndexOf("<div id=content-core>")).IndexOf("<div id=viewlet-below-content>"));

                                var listaExt = new List<string>();

                                /*Titulo*/

                                List<string> listTitulo = new List<string>() {  texto.IndexOf("<p align=center").ToString() + "|<p align=center|</p>"
                                                                               ,texto.IndexOf("<p class=1tituloacordo").ToString() + "|<p class=1tituloacordo|</p>"
                                                                               ,texto.IndexOf("<p class=a1-1tituloacordo").ToString() + "|<p class=a1-1tituloacordo|</p>"
                                                                               ,texto.IndexOf("<p class=11tituloacordo").ToString() + "|<p class=11tituloacordo|</p>"
                                                                               ,texto.IndexOf("<p class=a11tituloacordo").ToString() + "|<p class=a11tituloacordo|</p>"
                                                                               ,texto.IndexOf("<p class=tituloacordo").ToString() + "|<p class=tituloacordo|</p>"
                                                                               ,texto.IndexOf("<p class=a1-2tituloacordodp").ToString() + "|<p class=a1-2tituloacordodp|</p>"
                                                                               ,texto.IndexOf("<p class=12tituloacordodp").ToString() + "|<p class=12tituloacordodp|</p>"
                                                                               ,texto.IndexOf("<p class=textoacordo").ToString() + "|<p class=textoacordo|</p>"
                                                                               ,texto.IndexOf("<p class=msoheading9").ToString() + "|<p class=msoheading9|</p>"
                                                                               ,texto.IndexOf("<h2 class=11tituloacordo").ToString() + "|<h2 class=11tituloacordo|</h2>"
                                                                               ,texto.IndexOf("<title").ToString() + "|<title|</title>"
                                                                               ,texto.IndexOf("<p class=titulo").ToString() + "|<p class=titulo|</p>"
                                                                               ,texto.IndexOf("<p class=ttulotema").ToString() + "|<p class=ttulotema|</p>"};

                                listTitulo.RemoveAll(x => x.Split('|')[0].Contains("-1") || int.Parse(x.Split('|')[0]) > 1400);

                                listTitulo = listTitulo.OrderBy(x => Convert.ToInt32(x.Substring(0, x.IndexOf("|")))).ToList();

                                foreach (var itemLista in listTitulo)
                                {
                                    if (!itemLista.Contains("<p align=center") || (itemLista.Contains("<p align=center") && !texto.Contains("<p class=msotitle") && !texto.Contains("<p class=msobodytext ") && !texto.Contains("<p class=msonormal style")))
                                    {
                                        listaExt = Regex.Split(texto.ToLower(), itemLista.Split('|')[1]).ToList();

                                        listaExt.RemoveAt(0);
                                        listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                        listaExt.ForEach(x => titulo = titulo.Equals(string.Empty) ? ObterStringLimpa("<" + (x + itemLista.Split('|')[2]), itemLista.Split('|')[2]) : titulo);
                                    }

                                    if (titulo.Equals(string.Empty))
                                        continue;
                                    else
                                        break;
                                }

                                /*****Data Publicacao*****/

                                List<string> listDataPub = new List<string>() { texto.IndexOf("<p align=justify").ToString() + "|<p align=justify|<p>"
                                                                               ,texto.IndexOf("<p class=datapublicao").ToString() + "|<p class=datapublicao|</p>"
                                                                               ,texto.IndexOf("<p class=a2datapublicacao>").ToString() + "|<p class=a2datapublicacao>|</p>"
                                                                               ,texto.IndexOf("<p class=2datapublicao").ToString() + "|<p class=2datapublicao|</p>"
                                                                               ,texto.IndexOf("efeitos a partir de").ToString() + "|efeitos a partir de|</"
                                                                               ,texto.IndexOf("<b><b>,").ToString() + "|<b><b>,|</"
                                                                               ,texto.IndexOf("<li").ToString() + "|<li|</"
                                                                               ,texto.IndexOf("<span style=mso-bookmark:").ToString() + "|<span style=mso-bookmark:|</p>"
                                                                               ,texto.IndexOf("<p class=datadoe").ToString() + "|<p class=datadoe|</p>"
                                                                               ,texto.IndexOf("<p class=ttulotema").ToString() + "|<p class=ttulotema|</p>"};

                                listDataPub.RemoveAll(x => x.Contains("-1") || int.Parse(x.Split('|')[0]) > 1400);

                                listDataPub = listDataPub.OrderBy(x => Convert.ToInt32(x.Substring(0, x.IndexOf("|")))).ToList();

                                foreach (var itemLista in listDataPub)
                                {
                                    if (!itemLista.Contains("<p align=justify") || (itemLista.Contains("<p align=justify") && !texto.Contains("<p class=msotitle") && !texto.Contains("<p class=msobodytext ") && !texto.Contains("<p class=msonormal style")))
                                    {
                                        listaExt = Regex.Split(texto.ToLower(), itemLista.Split('|')[1]).ToList();

                                        listaExt.RemoveAt(0);
                                        listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty) || ObterStringLimpa("<" + x).Trim().Equals(titulo));
                                        listaExt.ForEach(x => publicacao = publicacao.Equals(string.Empty) ? ObterStringLimpa("<" + (x + itemLista.Split('|')[2]), itemLista.Split('|')[2]) : publicacao);
                                    }

                                    if (publicacao.Equals(string.Empty))
                                        continue;
                                    else
                                        break;
                                }

                                /******Novo Tratamento caso não encontre nos itens acima*****/

                                List<string> listString = new List<string>();

                                var textoAjustado = texto.Replace("<p></p>", string.Empty);

                                listString.Add(texto.IndexOf("<p class=msotitle").ToString() + "|<p class=msotitle|</p>");
                                listString.Add(texto.IndexOf("<p class=msobodytext ").ToString() + "|<p class=msobodytext |</p>");
                                listString.Add(texto.IndexOf("<h4").ToString() + "|<h4|</h4>");
                                listString.Add(texto.IndexOf("<h3").ToString() + "|<h3|</h3>");
                                listString.Add(texto.IndexOf("<h2").ToString() + "|<h2|</h2>");
                                listString.Add(texto.IndexOf("<h1").ToString() + "|<h1|</h1>");
                                listString.Add(texto.IndexOf("<p class=msonormal align=").ToString() + "|<p class=msonormal align=|</p>");
                                listString.Add(texto.IndexOf("<p class=msonormal style").ToString() + "|<p class=msonormal style|</p>");

                                listString.RemoveAll(x => x.Contains("-1") || int.Parse(x.Split('|')[0]) > 1400);

                                listString = listString.OrderBy(x => Convert.ToInt32(x.Substring(0, x.IndexOf("|")))).ToList();

                                int countItens = 0;

                                foreach (string item in listString)
                                {
                                    if (((texto.IndexOf("publicado") < texto.IndexOf("ato") && texto.IndexOf("publicado") > 0) ||
                                            (texto.IndexOf("publicação") < texto.IndexOf("ato") && texto.IndexOf("publicação") > 0) ||
                                                (texto.IndexOf("dou") < texto.IndexOf("ato") && texto.IndexOf("dou") > 0)) &&
                                                    urlTratada.ToLower().Contains("/atos"))
                                    {
                                        if (countItens == 0 && !publicacao.Equals(string.Empty) && listString.Count > 1)
                                        {
                                            countItens++;
                                            continue;
                                        }

                                        if (!item.Contains("<h3") && !item.Contains("<h2") && !item.Contains("<h1"))
                                        {
                                            listaExt = Regex.Split(textoAjustado.ToLower(), item.Split('|')[1]).ToList();

                                            listaExt.RemoveAt(0);
                                            listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty) || ObterStringLimpa("<" + x, "</p>").Trim().Length > 100);
                                            listaExt.ForEach(x => publicacao = publicacao.Equals(string.Empty) ||
                                                                    publicacao.Equals("ministério da fazenda") ||
                                                                        publicacao.Equals("comissão técnica permanente do icms-cotepe/icms") ||
                                                                            publicacao.Equals("secretaria-executiva") ||
                                                                                publicacao.Replace("–", "-").Equals("conselho nacional de política fazendária - confaz") ? ObterStringLimpa("<" + x, item.Split('|')[2]) : publicacao);
                                        }

                                        if (!item.Contains("<p class=msonormal") || listString[listString.Count - 1].Contains("<p class=msonormal"))
                                        {
                                            listaExt = Regex.Split(textoAjustado.ToLower(), item.Split('|')[1]).ToList();

                                            listaExt.RemoveAt(0);
                                            listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty) || publicacao.Equals(ObterStringLimpa("<" + x, item.Split('|')[2]).Trim()) || ObterStringLimpa("<" + x, item.Split('|')[2]).Trim().Length > 80);
                                            listaExt.ForEach(x => titulo = titulo.Equals(string.Empty) ||
                                                                    titulo.Equals("ministério da fazenda") ||
                                                                        titulo.Equals("comissão técnica permanente do icms-cotepe/icms") ||
                                                                            titulo.Equals("secretaria-executiva") ||
                                                                                titulo.Replace("–", "-").Equals("conselho nacional de política fazendária - confaz") ? ObterStringLimpa("<" + x, item.Split('|')[2]) : titulo);
                                        }
                                    }
                                    else
                                    {
                                        listaExt = Regex.Split(textoAjustado.ToLower(), item.Split('|')[1]).ToList();

                                        listaExt.RemoveAt(0);
                                        listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty) || ObterStringLimpa("<" + x, item.Split('|')[2]).Trim().Length > 80);
                                        listaExt.ForEach(x => titulo = titulo.Equals(string.Empty) ||
                                                                    titulo.Equals("ministério da fazenda") ||
                                                                        titulo.Equals("comissão técnica permanente do icms-cotepe/icms") ||
                                                                            titulo.Equals("secretaria-executiva") ||
                                                                                titulo.Replace("–", "-").Contains("conselho nacional de política fazendária - confaz") ? ObterStringLimpa("<" + x, item.Split('|')[2]) : titulo);

                                        listaExt = Regex.Split(textoAjustado.ToLower(), item.Split('|')[1]).ToList();

                                        listaExt.RemoveAt(0);
                                        listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty) || titulo.Equals(ObterStringLimpa("<" + x, item.Split('|')[2]).Trim()) || ObterStringLimpa("<" + x, item.Split('|')[2]).Trim().Length > 100 || (ObterStringLimpa("<" + x, item.Split('|')[2]).Trim().Equals(string.Empty) && ObterStringLimpa("<" + x).Trim().Length > 100));
                                        listaExt.ForEach(x => publicacao = publicacao.Equals(string.Empty) ? ObterStringLimpa("<" + x, item.Split('|')[2]) : publicacao);
                                    }

                                    countItens++;

                                    if ((!titulo.Equals(string.Empty) && !publicacao.Equals(string.Empty)) || countItens == 3)
                                        break;
                                    else
                                        continue;
                                }

                                if (titulo.Equals(string.Empty) && !publicacao.Equals(string.Empty))
                                    titulo = publicacao;

                                if (publicacao.Equals(string.Empty) || (!publicacao.Contains("dou") && !publicacao.Contains("publicado") && !publicacao.Contains("publicação") && !publicacao.Contains("efeitos")))
                                {
                                    int indexCaptura = 0;

                                    indexCaptura = texto.IndexOf("dou") > 0 ? texto.IndexOf("dou") : indexCaptura;
                                    indexCaptura = texto.IndexOf("publicado") > 0 && (texto.IndexOf("publicado") < indexCaptura || indexCaptura == 0) ? texto.IndexOf("publicado") : indexCaptura;
                                    indexCaptura = texto.IndexOf("publicação") > 0 && (texto.IndexOf("publicação") < indexCaptura || indexCaptura == 0) ? texto.IndexOf("publicação") : indexCaptura;

                                    publicacao = ObterStringLimpa(texto.Substring(indexCaptura).Substring(0, texto.Substring(indexCaptura).IndexOf("</p>")));
                                }

                                if (titulo.Equals(string.Empty) || publicacao.Equals(string.Empty))
                                    titulo = titulo;

                                contX = contX;
                                urlTratada = urlTratada;

                                hash = GerarHash(removerCaracterEspecial(removeTagScript(Regex.Replace(texto, @"\s+", " "))));

                                /*Ementa*/

                                listaExt = Regex.Split(texto.ToLower(), "<p class=a3ementa").ToList();

                                listaExt.RemoveAt(0);
                                listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                listaExt.ForEach(x => ementa = ementa.Equals(string.Empty) ? ObterStringLimpa("<" + x, "</p>") : ementa);

                                if (ementa.Equals(string.Empty))
                                {
                                    listaExt = Regex.Split(texto.ToLower(), "<p class=ementa").ToList();

                                    listaExt.RemoveAt(0);
                                    listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                    listaExt.ForEach(x => ementa = ementa.Equals(string.Empty) ? ObterStringLimpa("<" + x, "</p>") : ementa);
                                }

                                if (ementa.Equals(string.Empty) && texto.Contains("<p class=msotitle") && !texto.Contains("<p class=msosubtitle"))
                                {
                                    listaExt = Regex.Split(texto.ToLower(), "<p class=msobodytext").ToList();

                                    listaExt.RemoveAt(0);
                                    listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                    listaExt.ForEach(x => ementa = ementa.Equals(string.Empty) ? ObterStringLimpa("<" + x, "</p>") : ementa);
                                }

                                if (ementa.Equals(string.Empty) && texto.Contains("<p class=msotitle") && texto.Contains("<p class=msosubtitle"))
                                {
                                    listaExt = Regex.Split(texto.ToLower(), "<p class=msotitle").ToList();

                                    ementa = ObterStringLimpa("<" + listaExt[listaExt.Count - 1].Substring(0, listaExt[listaExt.Count - 1].IndexOf("</p>")));
                                }

                                if (ementa.Equals(string.Empty) && texto.Contains("<p class=msobodytext"))
                                {
                                    listaExt = Regex.Split(texto.ToLower(), "<p class=msobodytextindent3").ToList();

                                    listaExt.RemoveAt(0);
                                    listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                    listaExt.ForEach(x => ementa = ementa.Equals(string.Empty) ? ObterStringLimpa("<" + x, "</p>") : ementa);
                                }

                                if (ementa.Equals(string.Empty) && texto.Contains("<p class=msoblocktext style"))
                                {
                                    listaExt = Regex.Split(texto.ToLower(), "<p class=msoblocktext style").ToList();

                                    listaExt.RemoveAt(0);
                                    listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty));
                                    listaExt.ForEach(x => ementa = ementa.Equals(string.Empty) ? ObterStringLimpa("<" + x, "</p>") : ementa);
                                }

                                /*Captura CSS*/

                                var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                                listaCss.RemoveAt(0);

                                listaCssTratada = new List<string>();

                                string novaCss = string.Empty;

                                listaCss.ForEach(delegate(string x)
                                {
                                    novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                    if (!novaCss.ToLower().Contains("http"))
                                        novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                    listaCssTratada.Add(novaCss);
                                });

                                novaCss = string.Empty;

                                listaCssTratada.ForEach(x => novaCss += x);

                                string descStyle = string.Empty;

                                if (respostaNivel2_.Contains("<style"))
                                    descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                               .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                                /*Fim Captura CSS*/

                                texto = novaCss + descStyle + texto;
                            }

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                            string numero = string.Empty;
                            string especie = string.Empty;

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            //numero = titulo.Contains("/") ? titulo.Substring(0, titulo.IndexOf("/")).Trim() : titulo.Trim();

                            especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = Regex.Replace(texto, @"\s+", " ");
                            ementaInserir.Hash = hash;
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "FED";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > confaz > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz Bahia"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                webBrowser1 = new WebBrowser();

                webBrowser1.ScriptErrorsSuppressed = true;

                webBrowser1.Navigate("http://www.sefaz.ba.gov.br/motordebusca/pesquisa/Default.aspx?");

                webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted_SefazBa);
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazBa"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazBa");

                string urlTratada = string.Empty;

                List<string> xxxL = new List<string>();

                int contX = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (contX < 40)
                            {
                                contX++;
                                continue;
                            }

                            contX++;

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string nomeArq = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1);

                            string nomeForRep = nomeArq.Substring(0, nomeArq.LastIndexOf("."));

                            nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                            using (WebClient webclient = new WebClient())
                            {
                                webclient.DownloadFile(urlTratada, nomeArq);
                            }

                            string dadosPdf = LeArquivo(nomeArq);

                            byte[] arrayFile = File.ReadAllBytes(nomeArq);

                            ementaInserir.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = nomeArq.Substring(nomeArq.LastIndexOf(".") + 1), NomeArquivo = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1) });

                            File.Delete(nomeArq);

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "BA") + "}";

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = " ";

                            var listDados = Regex.Split(dadosPdf, "\n").ToList();

                            listDados.RemoveAll(x => x.Trim().Equals(string.Empty)
                                                        || x.ToLower().Trim().Equals("(revogada)")
                                                            || x.ToLower().Trim().Equals("(revogado)")
                                                                || x.ToLower().Trim().Equals("estado da bahia")
                                                                    || x.ToLower().Trim().Equals("secretaria da fazenda")
                                                                        || x.ToLower().Trim().Equals("governo do estado da bahia")
                                                                            || x.ToLower().Trim().Equals("conselho de fazenda estadual – consef")
                                                                                || x.ToLower().Trim().Equals("conselho de fazenda estadual")
                                                                                    || x.ToLower().Trim().Equals("secretaria da fazenda estadual")
                                                                                        || x.ToLower().Trim().Equals("conselho de fazenda do estado – consef")
                                                                                            || x.ToLower().Trim().Equals("secretraria da fazenda"));

                            if (listDados[0].Contains(nomeForRep))
                            {
                                nomeForRep = listDados[0];
                                listDados.RemoveAt(0);
                            }

                            titulo = listDados[0];

                            var posicaoRep = new List<int>() { titulo.ToLower().IndexOf("de "), titulo.IndexOf(",") }.OrderBy(x => x).ToList();
                            posicaoRep.RemoveAll(x => x == -1);

                            if (posicaoRep.Count > 0)
                                edicao = titulo.Substring(posicaoRep[0] + 2).Trim();

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                            publicacao = listDados[1];

                            foreach (var x in listDados.GetRange(2, listDados.Count - 2))
                            {
                                if (!x.ToLower().Contains("o secretário da fazenda"))
                                    ementa += "\n" + x;
                                else
                                    break;
                            }

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = dadosPdf.Replace(nomeForRep, string.Empty);
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(dadosPdf));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "BA";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > sefazBa > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz ES"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                var listUrlEs = obterUrlNivelDocumento("http://www.sefaz.es.gov.br/LegislacaoOnline/lpext.dll?f=templates&fn=tools-contents.htm&cp=InfobaseLegislacaoOnline&2.0");
            }

            #endregion

            #region "Captura DOC's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("d") && siglaFonteProcessamento.Equals("sefazEs"))
            {
                string hash = string.Empty;
                string texto = string.Empty;
                string titulo = string.Empty;
                string publicacao = string.Empty;
                string ementa = string.Empty;
                string urlTratada = string.Empty;

                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazEs");

                var listaExt = new List<string>();

                dynamic ementaInserir;
                int contX = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (contX < 0 || itemLista_Nivel2.Url.Trim().Contains(".jpg"))
                            {
                                contX++;
                                continue;
                            }

                            contX++;

                            Thread.Sleep(2000);

                            hash = string.Empty;
                            texto = string.Empty;
                            titulo = string.Empty;
                            publicacao = string.Empty;
                            ementa = string.Empty;

                            ementaInserir = new ExpandoObject();
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            WebRequest reqNivel2_ = WebRequest.Create(urlTratada);
                            WebResponse resNivel2_ = reqNivel2_.GetResponse();
                            Stream dataStreamNivel2_ = resNivel2_.GetResponseStream();
                            StreamReader readerNivel2_;

                            readerNivel2_ = new StreamReader(dataStreamNivel2_, Encoding.Default);

                            string respostaNivel2_ = readerNivel2_.ReadToEnd();
                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", string.Empty).Replace("\r", string.Empty).Replace("\t", string.Empty));

                            texto = respostaNivel2_.ToLower().Substring(respostaNivel2_.ToLower().LastIndexOf("<body")).Substring(0, respostaNivel2_.ToLower().Substring(respostaNivel2_.ToLower().LastIndexOf("<body")).IndexOf("</body>"));

                            List<string> listString = new List<string>();

                            var textoAjustado = texto.Replace("<p></p>", string.Empty);

                            listString.Add(texto.IndexOf("<p class=msotitle").ToString() + "|<p class=msotitle|</p>|");
                            listString.Add(texto.IndexOf("<p class=msobodytext ").ToString() + "|<p class=msobodytext |</p>|");
                            listString.Add(texto.IndexOf("<p class=msonormal ").ToString() + "|<p class=msonormal |</p>|");
                            listString.Add(texto.IndexOf("<p class=msonormal>").ToString() + "|<p class=msonormal>|</p>|1");
                            listString.Add(texto.IndexOf("<div style=border:solid windowtext").ToString() + "|<div style=border:solid windowtext|</div>|T");
                            listString.Add(texto.IndexOf("<p class=a010165").ToString() + "|<p class=a010165|</p>|T");
                            listString.Add(texto.IndexOf("<p class=msoblocktext").ToString() + "|<p class=msoblocktext|</p>|");
                            listString.Add(texto.IndexOf("<p class=msoplaintext").ToString() + "|<p class=msoplaintext|</p>|");
                            listString.Add(texto.IndexOf("<p align=center").ToString() + "|<p align=center|</p>|");
                            listString.Add(texto.IndexOf("<p class=bodytext30").ToString() + "|<p class=bodytext30|</p>|");
                            listString.Add(texto.IndexOf("<p class=msobodytextindent").ToString() + "|<p class=msobodytextindent|</p>|E");
                            listString.Add(texto.IndexOf("<p align=justify").ToString() + "|<p align=justify|</p>|");
                            listString.Add(texto.IndexOf("<p class=blockquote").ToString() + "|<p class=blockquote|</p>|");
                            listString.Add(texto.IndexOf("<h1").ToString() + "|<h1|</h1>|");
                            listString.Add(texto.IndexOf("<h2").ToString() + "|<h2|</h2>|");
                            listString.Add(texto.IndexOf("<h4 align=center").ToString() + "|<h4 align=center|</h4>|");
                            listString.Add(texto.IndexOf("<h6 align=center").ToString() + "|<h6 align=center|</h6>|");
                            listString.Add(texto.IndexOf("<h3 align=center").ToString() + "|<h3 align=center|</h3>|");
                            listString.Add(texto.IndexOf("<h5 align=center").ToString() + "|<h5 align=center|</h5>|");
                            listString.Add(texto.IndexOf("<p class=msocaption").ToString() + "|<p class=msocaption|</p>|");
                            listString.Add(texto.IndexOf("<p class=msoheading8").ToString() + "|<p class=msoheading8|</p>|");
                            listString.Add(texto.IndexOf("<p class=padrao").ToString() + "|<p class=padrao|</p>|");
                            listString.Add(texto.IndexOf("<b><p>").ToString() + "|<b><p>|</p>|1");
                            listString.Add(texto.IndexOf("<font face=arial><p>").ToString() + "|<font face=arial><p>|</p>|1");
                            listString.Add(texto.IndexOf("<p class=bodytext30").ToString() + "|<p class=bodytext30|</p>|1");

                            listString.RemoveAll(x => x.Contains("-1"));

                            listString = listString.OrderBy(x => Convert.ToInt32(x.Substring(0, x.IndexOf("|")))).ToList();

                            int indexCorte = 0;

                            if (listString.Count > 2)
                                indexCorte = int.Parse(listString[1].Split('|')[0]);

                            else if (listString.Count > 0)
                                indexCorte = int.Parse(listString[listString.Count - 1].Split('|')[0]);

                            listString.RemoveAll(x => (int.Parse(x.Split('|')[0]) > indexCorte && !x.Contains("|E")));

                            int countItens = 0;

                            string regOrdem = respostaNivel2_.ToLower().Substring(respostaNivel2_.ToLower().IndexOf("<title") + 7).Substring(0, respostaNivel2_.ToLower().Substring(respostaNivel2_.ToLower().IndexOf("<title") + 7).IndexOf("</title>")).Trim();

                            if (texto.IndexOf("<title") >= 0)
                                texto = texto.Remove(texto.IndexOf("<title"), texto.IndexOf("</title>") - texto.IndexOf("<title"));

                            else if (urlTratada.Contains("ordens%20de%20servi%e7o"))
                                regOrdem = "ordem";

                            if (regOrdem.IndexOf(" ") > 0)
                                regOrdem = regOrdem.Substring(0, regOrdem.IndexOf(" ")).ToLower();

                            foreach (string item in listString)
                            {
                                if (!item.Contains("|E"))
                                {
                                    if ((texto.IndexOf("doe") < texto.IndexOf(regOrdem) && texto.IndexOf("doe") > 0) ||
                                            (texto.IndexOf("d.o") < texto.IndexOf(regOrdem) && texto.IndexOf("d.o") > 0) ||
                                                (texto.IndexOf("d.º") < texto.IndexOf(regOrdem) && texto.IndexOf("d.º") > 0) ||
                                                    (texto.IndexOf("do:") < texto.IndexOf(regOrdem) && texto.IndexOf("do:") > 0) ||
                                                        (texto.IndexOf("dio") < texto.IndexOf(regOrdem) && texto.IndexOf("dio") > 0) ||
                                                              (texto.IndexOf("d.i.o") < texto.IndexOf(regOrdem) && texto.IndexOf("d.i.o") > 0) ||
                                                                    (texto.IndexOf("d. o . e") < texto.IndexOf(regOrdem) && texto.IndexOf("d. o . e") > 0) ||
                                                                        (texto.IndexOf("dou") < texto.IndexOf(regOrdem) && texto.IndexOf("dou") > 0))
                                    {
                                        if (countItens == 0 && !publicacao.Equals(string.Empty) && listString.Count > 1)
                                        {
                                            countItens++;
                                            continue;
                                        }

                                        listaExt = Regex.Split(textoAjustado.ToLower(), item.Split('|')[1]).ToList();

                                        listaExt.RemoveAt(0);
                                        listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty) ||
                                                                    ObterStringLimpa("<" + x, "</p>").Trim().Length > 100 ||
                                                                        (!ObterStringLimpa("<" + x, "</p>").Contains("doe") &&
                                                                            !ObterStringLimpa("<" + x, "</p>").Contains("d.o") &&
                                                                                !ObterStringLimpa("<" + x, "</p>").Contains("d.º") &&
                                                                                    !ObterStringLimpa("<" + x, "</p>").Contains("do:") &&
                                                                                        !ObterStringLimpa("<" + x, "</p>").Contains("dio") &&
                                                                                            !ObterStringLimpa("<" + x, "</p>").Contains("d.i.o") &&
                                                                                                !ObterStringLimpa("<" + x, "</p>").Contains("d. o . e") &&
                                                                                                    !ObterStringLimpa("<" + x, "</p>").Contains("dou")));

                                        listaExt.ForEach(x => publicacao = publicacao.Equals(string.Empty) ||
                                                                publicacao.Equals("governo do estado do espírito santo secretaria de estado da fazenda") ||
                                                                    publicacao.Equals("governo do estado do espírito santo") ||
                                                                        publicacao.Equals("minuta") ? ObterStringLimpa("<" + x, item.Split('|')[2]) : publicacao);

                                        listaExt = Regex.Split(textoAjustado.ToLower(), item.Split('|')[1]).ToList().Select(x => textoAjustado.IndexOf(x) + "|" + x).ToList();

                                        listaExt.RemoveAt(0);
                                        int indexMin = textoAjustado.IndexOf(publicacao);

                                        listaExt = listaExt.Where(y => int.Parse(y.Split('|')[0]) < (indexCorte + 500) && int.Parse(y.Split('|')[0]) > indexMin).ToList();

                                        listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty) || publicacao.Equals(ObterStringLimpa("<" + x, item.Split('|')[2]).Trim()) || ObterStringLimpa("<" + x, item.Split('|')[2]).Trim().Length > 80);
                                        listaExt.ForEach(x => titulo = titulo.Equals(string.Empty) ||
                                                                titulo.Equals("governo do estado do espírito santo secretaria de estado da fazenda") ||
                                                                    titulo.Equals("governo do estado do espírito santo") ||
                                                                        titulo.Equals("minuta") ? ObterStringLimpa("<" + x, item.Split('|')[2]) : titulo);
                                    }
                                    else
                                    {
                                        listaExt = Regex.Split(textoAjustado.ToLower(), item.Split('|')[1]).ToList();

                                        string compTag = item.Split('|')[3].Equals("1") ? string.Empty : "<";

                                        listaExt.RemoveAt(0);
                                        listaExt.RemoveAll(x => ObterStringLimpa(compTag + x).Trim().Equals(string.Empty) || ObterStringLimpa(compTag + x, item.Split('|')[2]).Trim().Length > 100 || !x.Contains(regOrdem));
                                        listaExt.ForEach(x => titulo = titulo.Equals(string.Empty) ||
                                                                    titulo.Equals("governo do estado do espírito santo secretaria de estado da fazenda") ||
                                                                        titulo.Equals("governo do estado do espírito santo") ||
                                                                            titulo.Equals("minuta") ? ObterStringLimpa(compTag + x, item.Split('|')[2]) : titulo);

                                        listaExt = Regex.Split(textoAjustado.ToLower(), item.Split('|')[1]).ToList();

                                        listaExt.RemoveAt(0);
                                        listaExt.RemoveAll(x => ObterStringLimpa(compTag + x).Trim().Equals(string.Empty) || titulo.Equals(ObterStringLimpa(compTag + x, item.Split('|')[2]).Trim()) || ObterStringLimpa(compTag + x, item.Split('|')[2]).Trim().Length > 100 || (ObterStringLimpa(compTag + x, item.Split('|')[2]).Trim().Equals(string.Empty) && ObterStringLimpa(compTag + x).Trim().Length > 100));
                                        listaExt.ForEach(x => publicacao = publicacao.Equals(string.Empty) ||
                                                                    publicacao.Equals("governo do estado do espírito santo secretaria de estado da fazenda") ||
                                                                        publicacao.Equals("governo do estado do espírito santo") ||
                                                                            publicacao.Equals("minuta") ? ObterStringLimpa(compTag + x, item.Split('|')[2]) : publicacao);
                                    }
                                }

                                listaExt = Regex.Split(textoAjustado.ToLower(), item.Split('|')[1]).ToList();

                                listaExt.RemoveAt(0);
                                listaExt.RemoveAll(x => ObterStringLimpa("<" + x).Trim().Equals(string.Empty) ||
                                                            publicacao.Equals(ObterStringLimpa("<" + x, item.Split('|')[2]).Trim()) ||
                                                               titulo.Equals(ObterStringLimpa("<" + x, item.Split('|')[2]).Trim()));

                                listaExt.ForEach(x => ementa = ementa.Equals(string.Empty) ||
                                                        ementa.Equals("governo do estado do espírito santo secretaria de estado da fazenda") ||
                                                            ementa.Equals("governo do estado do espírito santo") ||
                                                                ementa.Equals("minuta") ? ObterStringLimpa("<" + x, item.Split('|')[2]) : ementa);

                                countItens++;

                                if ((!titulo.Equals(string.Empty) && !publicacao.Equals(string.Empty) && !ementa.Equals(string.Empty)) || countItens == 3)
                                    break;
                                else
                                    continue;
                            }

                            if (publicacao.Equals(string.Empty) ||
                                    (!publicacao.Contains("d.o") &&
                                        !publicacao.Contains("doe") &&
                                            !publicacao.Contains("d.º") &&
                                                !publicacao.Contains("do:") &&
                                                    !publicacao.Contains("efeitos a partir de ") &&
                                                        !publicacao.Contains("dio") &&
                                                            !publicacao.Contains("d.i.o") &&
                                                                !publicacao.Contains("d. o . e") &&
                                                                    !publicacao.Contains("dou")))
                            {
                                int indexCaptura = 0;
                                indexCaptura = texto.IndexOf("d.o");
                                indexCaptura = texto.IndexOf("doe") > 0 && (texto.IndexOf("doe") < indexCaptura || indexCaptura <= 0) ? texto.IndexOf("doe") : indexCaptura;
                                indexCaptura = texto.IndexOf("d.º") > 0 && (texto.IndexOf("d.º") < indexCaptura || indexCaptura <= 0) ? texto.IndexOf("d.º") : indexCaptura;
                                indexCaptura = texto.IndexOf("do:") > 0 && (texto.IndexOf("do:") < indexCaptura || indexCaptura <= 0) ? texto.IndexOf("do:") : indexCaptura;
                                indexCaptura = texto.IndexOf("dio") > 0 && (texto.IndexOf("dio") < indexCaptura || indexCaptura <= 0) ? texto.IndexOf("dio") : indexCaptura;
                                indexCaptura = texto.IndexOf("d.i.o") > 0 && (texto.IndexOf("d.i.o") < indexCaptura || indexCaptura <= 0) ? texto.IndexOf("d.i.o") : indexCaptura;
                                indexCaptura = texto.IndexOf("dou") > 0 && (texto.IndexOf("dou") < indexCaptura || indexCaptura <= 0) ? texto.IndexOf("dou") : indexCaptura;
                                indexCaptura = texto.IndexOf("d. o . e") > 0 && (texto.IndexOf("d. o . e") < indexCaptura || indexCaptura <= 0) ? texto.IndexOf("d. o . e") : indexCaptura;
                                indexCaptura = texto.IndexOf("efeitos a partir de") > 0 && (texto.IndexOf("efeitos a partir de") < indexCaptura || indexCaptura <= 0) ? texto.IndexOf("efeitos a partir de") : indexCaptura;

                                if (indexCaptura > 0)
                                    publicacao = ObterStringLimpa(texto.Substring(indexCaptura).Substring(0, texto.Substring(indexCaptura).IndexOf("</p>")));
                                else
                                    publicacao = string.Empty;
                            }

                            if (titulo.Equals(string.Empty) || publicacao.Equals(string.Empty))
                                titulo = titulo;

                            contX = contX;
                            urlTratada = urlTratada;

                            hash = GerarHash(removerCaracterEspecial(removeTagScript(Regex.Replace(texto, @"\s+", " "))));

                            /*Captura CSS*/

                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (respostaNivel2_.Contains("<style"))
                                descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                           .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            texto = novaCss + descStyle + texto;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "ES") + "}";

                            string numero = string.Empty;
                            string especie = string.Empty;

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = Regex.Replace(texto, @"\s+", " ");
                            ementaInserir.Hash = hash;
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "ES";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > confaz > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz DF"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                List<string> listUrlDF = new List<string>() { "L412", "L452", "L151", "L150", "L573", "L170", "L12", "L90", "L3", "L4", "L290", "L9", "L432", "L6", "L8", "L5", "L7" };

                string prefixUrl = "http://www.fazenda.df.gov.br/aplicacoes/legislacao/legislacao/Pes_Av_Legislacao.cfm?FromRec=1&NumeroResultados=200&txtTodasP=&txtExpressao=&txtQualquerP=&txtTrechoP=&txtSemP=&rdbLocal=doc&chkTodas=&chkTipo={0}&chkNumero=&txtNumero=&chkAno=&txtAno=&chkData=periodo&txtDataIni={1}&txtDataFim={2}&rdbOrdena=Hierarquica&rdbOrdenaData=decrescente&paginacao=true";
                string resposta = string.Empty;

                try
                {
                    foreach (var itemUrlDF in listUrlDF)
                    {
                        objUrll = new ExpandoObject();

                        objUrll.Indexacao = "Secretaria do estado da fazenda do DF";
                        objUrll.Url = string.Format(prefixUrl, itemUrlDF, string.Empty, string.Empty);
                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        resposta = getHtmlPaginaByGet(objUrll.Url, string.Empty);

                        if (resposta.ToLower().Contains("refine mais sua pesquisa!"))
                        {
                            int intervalo = 30;
                            DateTime dataInicial = new DateTime(1960, 01, 01);

                            while (dataInicial.Year <= DateTime.Now.Year)
                            {
                                string urlFormatada = string.Format(prefixUrl, itemUrlDF, dataInicial.ToString("dd/MM/yyyy"), dataInicial.AddYears(intervalo).ToString("dd/MM/yyyy"));

                                resposta = getHtmlPaginaByGet(urlFormatada, string.Empty);

                                if (resposta.ToLower().Contains("refine mais sua pesquisa!"))
                                    intervalo--;

                                else
                                {
                                    if (dataInicial.Equals(new DateTime(1960, 01, 01)))
                                    {
                                        dataInicial = dataInicial.AddYears(30);
                                        intervalo = 5;
                                    }
                                    else
                                        dataInicial = dataInicial.AddYears(intervalo);

                                    if (intervalo < 5)
                                        intervalo++;

                                    //tratamento para pegar os results.
                                    resposta = resposta.Replace("\"", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty).ToLower();
                                    resposta = resposta.Substring(resposta.IndexOf("<table class=conteudoarea width=")).Substring(0, resposta.Substring(resposta.IndexOf("<table class=conteudoarea width=")).IndexOf("</table>"));
                                    var listaResults = Regex.Split(resposta, "href=").ToList();

                                    listaResults.RemoveAt(0);
                                    listaResults.ForEach(delegate(string x)
                                    {
                                        try
                                        {
                                            objItenUrl = new ExpandoObject();
                                            objItenUrl.Url = x.Substring(0, x.IndexOf(" "));
                                            objUrll.Lista_Nivel2.Add(objItenUrl);
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    });
                                }
                            }
                        }
                        else
                        {
                            //tratamento para pegar os results.
                            resposta = resposta.Replace("\"", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty).ToLower();
                            resposta = resposta.Substring(resposta.IndexOf("<table class=conteudoarea width=")).Substring(0, resposta.Substring(resposta.IndexOf("<table class=conteudoarea width=")).IndexOf("</table>"));
                            var listaResults = Regex.Split(resposta, "href=").ToList();

                            listaResults.RemoveAt(0);
                            listaResults.ForEach(delegate(string x)
                            {
                                try
                                {
                                    objItenUrl = new ExpandoObject();
                                    objItenUrl.Url = x.Substring(0, x.IndexOf(" "));
                                    objUrll.Lista_Nivel2.Add(objItenUrl);
                                }
                                catch (Exception ex)
                                {
                                }
                            });
                        }

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                }
                catch (Exception ex)
                {
                }
            }

            #endregion

            #region "Captura DOC's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("d") && siglaFonteProcessamento.Equals("sefazDf"))
            {
                string hash = string.Empty;
                string texto = string.Empty;
                string titulo = string.Empty;
                string publicacao = string.Empty;
                string ementa = string.Empty;
                string urlTratada = string.Empty;

                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazDf");

                var listaExt = new List<string>();

                dynamic ementaInserir;
                int contX = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (/*!itemLista_Nivel2.Url.Trim().Equals("http://www.fazenda.df.gov.br//aplicacoes/legislacao/legislacao/telasaidadocumento.cfm?txtnumero=4&txtano=1994&txttipo=4&txtparte=.")*/contX < 572)
                            {
                                contX++;
                                continue;
                            }

                            contX++;

                            Thread.Sleep(2000);

                            hash = string.Empty;
                            texto = string.Empty;
                            titulo = string.Empty;
                            publicacao = string.Empty;
                            ementa = string.Empty;

                            ementaInserir = new ExpandoObject();
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, string.Empty);
                            List<string> listString = new List<string>();

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " ")).ToLower();

                            listString.Add(respostaNivel2_.IndexOf("<div class=section1") + "|<div class=section1");
                            listString.Add(respostaNivel2_.IndexOf("<div class=wordsection1") + "|<div class=wordsection1");

                            listString.RemoveAll(x => x.Contains("-1"));
                            listString = listString.OrderBy(x => Convert.ToInt32(x.Substring(0, x.IndexOf("|")))).ToList();

                            if (listString.Count == 0)
                                continue;

                            texto = respostaNivel2_.Substring(respostaNivel2_.IndexOf(listString[0].Split('|')[1])).Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf(listString[0].Split('|')[1])).IndexOf("</body>"));

                            string keyTitulo = "";

                            if (respostaNivel2_.IndexOf("<title") > 0)
                            {
                                keyTitulo = ObterStringLimpa(respostaNivel2_.Substring(respostaNivel2_.IndexOf("<title")).Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<title")).IndexOf("</title>"))).Trim().Split(' ')[0];

                                keyTitulo = keyTitulo.Split(' ')[0] + " ";

                                if (texto.IndexOf(keyTitulo) > 0)
                                    titulo = texto.Substring(texto.IndexOf(keyTitulo)).Substring(0, texto.Substring(texto.IndexOf(keyTitulo)).IndexOf("</p>"));
                                else
                                    titulo = ObterStringLimpa(texto.Substring(0, texto.IndexOf("</p>")));
                            }
                            else
                                titulo = ObterStringLimpa(texto.Substring(0, texto.IndexOf("</p>")));

                            //keyTitulo = keyTitulo.Split(' ')[0] + " ";
                            //titulo = texto.Substring(texto.IndexOf(keyTitulo)).Substring(0, texto.Substring(texto.IndexOf(keyTitulo)).IndexOf("</p>"));

                            titulo = ObterStringLimpa(titulo);

                            listString = new List<string>();

                            listString.Add(texto.IndexOf("publicação") + "|publicação");
                            listString.Add(texto.IndexOf("publicado") + "|publicado");
                            listString.Add(texto.IndexOf("publicada") + "|publicada");

                            listString.RemoveAll(x => x.Contains("-1") || (int.Parse(x.Split('|')[0]) > texto.IndexOf(keyTitulo) + 1000 && texto.IndexOf(keyTitulo) > 0));
                            listString = listString.OrderBy(x => Convert.ToInt32(x.Substring(0, x.IndexOf("|")))).ToList();

                            if (listString.Count > 0)
                                publicacao = texto.Substring(texto.IndexOf(listString[0].Split('|')[1])).Substring(0, texto.Substring(texto.IndexOf(listString[0].Split('|')[1])).IndexOf("</p>"));

                            publicacao = ObterStringLimpa(publicacao);

                            if (titulo.Equals(string.Empty) || publicacao.Equals(string.Empty))
                                titulo = titulo;

                            contX = contX;
                            urlTratada = urlTratada;

                            hash = GerarHash(removerCaracterEspecial(removeTagScript(Regex.Replace(texto, @"\s+", " "))));

                            /*Captura CSS*/

                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (respostaNivel2_.Contains("<style"))
                                descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                           .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            texto = novaCss + descStyle + texto;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "DF") + "}";

                            string numero = string.Empty;
                            string especie = string.Empty;

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = Regex.Replace(texto, @"\s+", " ");
                            ementaInserir.Hash = hash;
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "DF";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > sefazEs > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz MG"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                var listSefazMG = new List<string>() { "http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/comunicados/"
                                                      ,"http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/portarias/"
                                                      ,"http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/resolucoes/"
                                                      ,"http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/decretos/"
                                                      ,"http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/leis/"
                                                      ,"http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/instrucoes_normativas/"
                                                      ,"http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/downloads/"
                                                      ,"http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/ricms_2002_seco/sumario2002seco.htm"
                                                      ,"http://www.almg.gov.br/export/sites/default/consulte/legislacao/Downloads/pdfs/constituicao_estadual_multivigente.pdf"
                                                      ,"http://www.almg.gov.br/export/sites/default/consulte/legislacao/Downloads/pdfs/ConstituicaoEstadual.pdf" };

                foreach (var itemSefazMG in listSefazMG)
                {
                    var listUrlMg = new List<string>();

                    if (!itemSefazMG.Contains(".pdf"))
                        listUrlMg = obterUrlNivelDocumentoSefazMG(itemSefazMG, itemSefazMG);

                    else
                        listUrlMg = new List<string>() { itemSefazMG };

                    objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Secretaria do estado da fazenda de MG";
                    objUrll.Url = itemSefazMG;
                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    listUrlMg.ForEach(delegate(string itemDado)
                    {
                        objItenUrl = new ExpandoObject();

                        objItenUrl.Url = itemDado;
                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    });

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                }
            }

            #endregion

            #region "Captura DOC's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("d") && siglaFonteProcessamento.Equals("sefazMg"))
            {
                string hash = string.Empty;
                string texto = string.Empty;
                string titulo = string.Empty;
                string publicacao = string.Empty;
                string ementa = string.Empty;
                string urlTratada = string.Empty;

                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazMg");

                var listaExt = new List<string>();

                dynamic ementaInserir;
                int contX = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (contX < 0)
                            {
                                contX++;
                                continue;
                            }

                            contX++;

                            Thread.Sleep(2000);

                            hash = string.Empty;
                            texto = string.Empty;
                            titulo = string.Empty;
                            publicacao = string.Empty;
                            ementa = string.Empty;

                            ementaInserir = new ExpandoObject();
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, "default");
                            List<string> listString = new List<string>();

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", "").Replace("\n", " ").Replace("\r", " ").Replace("\t", " ")).ToLower();

                            listString.Add(respostaNivel2_.IndexOf("<div class=wordsection1") + "|<div class=wordsection1");
                            listString.Add(respostaNivel2_.IndexOf("<td style=width:70%") + "|<td style=width:70%");

                            listString.RemoveAll(x => x.Contains("-1"));

                            if (listString.Count == 0)
                                continue;

                            texto = respostaNivel2_.Substring(respostaNivel2_.IndexOf(listString[0].Split('|')[1])).Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf(listString[0].Split('|')[1])).IndexOf("</body>"));

                            string keyTitulo = "";

                            if (respostaNivel2_.IndexOf("<title") > 0)
                            {
                                keyTitulo = ObterStringLimpa(respostaNivel2_.Substring(respostaNivel2_.IndexOf("<title")).Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<title")).IndexOf("</title>"))).Trim().Split(' ')[0];

                                keyTitulo = keyTitulo.Split(' ')[0] + " ";

                                if (texto.IndexOf(keyTitulo) > 0)
                                    titulo = texto.Substring(texto.IndexOf(keyTitulo)).Substring(0, texto.Substring(texto.IndexOf(keyTitulo)).IndexOf("</p>"));
                                else
                                    titulo = ObterStringLimpa(texto.Substring(0, texto.IndexOf("</p>")));
                            }
                            else
                                titulo = ObterStringLimpa(texto.Substring(0, texto.IndexOf("</p>")));

                            titulo = ObterStringLimpa(titulo);

                            publicacao = titulo.IndexOf("(") > 0 ? titulo.Substring(titulo.IndexOf("(")) : string.Empty;

                            if (!publicacao.Equals(string.Empty))
                                titulo = titulo.Replace(publicacao, string.Empty);
                            else
                            {
                                var listPub = new List<string>();

                                listPub.Add(texto.IndexOf("(pub.").ToString() + "|(pub.");
                                listPub.Add(texto.IndexOf("(mg").ToString() + "|(mg");

                                listPub.RemoveAll(x => x.Contains("-1"));

                                if (listPub.Count > 0)
                                    publicacao = texto.Substring(texto.IndexOf(listPub[0].Split('|')[1])).Substring(0, texto.Substring(texto.IndexOf(listPub[0].Split('|')[1])).IndexOf("</p>"));
                            }

                            if (titulo.Equals(string.Empty) || publicacao.Equals(string.Empty))
                                titulo = titulo;

                            contX = contX;
                            urlTratada = urlTratada;

                            hash = GerarHash(removerCaracterEspecial(removeTagScript(Regex.Replace(texto, @"\s+", " "))));

                            /*Captura CSS*/

                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (respostaNivel2_.Contains("<style"))
                                descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                           .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            texto = novaCss + descStyle + texto;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "MG") + "}";

                            string numero = string.Empty;
                            string especie = string.Empty;

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = Regex.Replace(texto, @"\s+", " ");
                            ementaInserir.Hash = hash;
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "MG";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > sefazMG > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz RJ"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                var listSefazRJ = new List<string>() { /*"http://alerjln1.alerj.rj.gov.br/constest.nsf/indiceint?openform|N"
                                                      ,"http://alerjln1.alerj.rj.gov.br/contlei.nsf/leicompint?openform|N"
                                                      ,"http://alerjln1.alerj.rj.gov.br/contlei.nsf/leiordint?openform|N"*/
                                                       "http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna1/menu_legislacao_decretos/Decretos-Tributaria|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna1/menu_legislacao_decretos/Decretos-Financeira|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna2/RegulamentoDoICMS|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna2/menu_legislacao_resolucoes/Resolucoes-Financeira|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna2/menu_legislacao_resolucoes/Resolucoes-Tributaria|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna2/menu_legislacao_resolucoes-conjuntas/ResolucoesConjuntas-Financeira|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna2/menu_legislacao_resolucoes-conjuntas/ResolucoesConjuntas-Tributaria|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna3/Portarias/Portarias-Administrativa|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna3/Portarias/Portarias-Financeira|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna3/Portarias/Portarias-Tributaria|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna3/PortariasConjuntas/PortariasConjuntas-Financeira|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna3/PortariasConjuntas/PortariasConjuntas-Tributaria|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/folder73/menu_legislacao_instrucoesnormativas/InstrucoesNormativas-Administrativa|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/folder73/menu_legislacao_instrucoesnormativas/InstrucoesNormativas-Financeira|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/folder73/menu_legislacao_instrucoesnormativas/InstrucoesNormativas-Tributaria|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/folder73/url70|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/folder73/url70|s"
                                                      ,"http://www.fazenda.rj.gov.br/sefaz/faces/menu_structure/legislacao/legislacao-estadual-navigation/coluna2/RegulamentoDoICMS|s"};

                foreach (var itemSefazRJ in listSefazRJ)
                {
                    var listUrlMg = new List<string>();

                    if (!itemSefazRJ.Contains(".pdf"))
                        listUrlMg = obterUrlNivelDocumentoSefazRJ(itemSefazRJ, itemSefazRJ);

                    else
                        listUrlMg = new List<string>() { itemSefazRJ };

                    objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Secretaria do estado da fazenda de RJ";
                    objUrll.Url = itemSefazRJ.Replace("|N", string.Empty).Replace("|s", string.Empty);
                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    listUrlMg.ForEach(delegate(string itemDado)
                    {
                        objItenUrl = new ExpandoObject();

                        objItenUrl.Url = itemDado;
                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    });

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                }
            }

            #endregion

            #region "Captura DOC's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("d") && siglaFonteProcessamento.Equals("sefazRj"))
            {
                string hash = string.Empty;
                string texto = string.Empty;
                string titulo = string.Empty;
                string publicacao = string.Empty;
                string ementa = string.Empty;
                string urlTratada = string.Empty;

                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazRj");

                var listaExt = new List<string>();
                dynamic ementaInserir;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            hash = string.Empty;
                            texto = string.Empty;
                            titulo = string.Empty;
                            publicacao = string.Empty;
                            ementa = string.Empty;

                            ementaInserir = new ExpandoObject();
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, "default");
                            List<string> listString = new List<string>();

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", string.Empty).Replace("\n", " ").Replace("\r", " ").Replace("\t", " ")).ToLower();

                            listString.Add(respostaNivel2_.IndexOf("<body") + "|<body");

                            texto = respostaNivel2_.Substring(respostaNivel2_.IndexOf(listString[0].Split('|')[1])).Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf(listString[0].Split('|')[1])).IndexOf("=#topo>"));

                            if (!urlTratada.Contains("constest.nsf/") && texto.Contains("<font face=arial>lei"))
                            {
                                titulo = texto.Substring(texto.IndexOf("<font face=arial>lei")).Substring(0, texto.Substring(texto.IndexOf("<font face=arial>lei")).IndexOf("</font>"));

                                titulo = ObterStringLimpa(titulo);

                                publicacao = titulo.Substring(titulo.IndexOf(",") + 1);
                            }
                            else
                                continue;

                            hash = GerarHash(removerCaracterEspecial(removeTagScript(Regex.Replace(texto, @"\s+", " "))));

                            /*Captura CSS*/

                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (respostaNivel2_.Contains("<style"))
                                descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                           .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                            /**Fim Captura CSS**/
                            texto = novaCss + descStyle + texto;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "RJ") + "}";

                            string numero = string.Empty;
                            string especie = string.Empty;

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = Regex.Replace(texto, @"\s+", " ");
                            ementaInserir.Hash = hash;
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "RJ";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > sefazRJ > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz SP"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                this.Text = "Fonte: Sefaz SP";

                var listaUrlsSP = new List<string>(){"http://info.fazenda.sp.gov.br/NXT/gateway.dll/legislacao_tributaria/agendas/legistrib2.htm"/*,
                                                     "http://info.fazenda.sp.gov.br/NXT/gateway.dll/?f=templates$fn=document-frame.htm$3.0"*/};

                foreach (var itemSp in listaUrlsSP)
                {
                    WebRequest req = WebRequest.Create(itemSp);
                    req.Headers.Set("Cookie", "nxtopts=420zhlf0zzzzz; nxt/gateway.dll/uid=36E76821; nxt/gateway.dll/sid=36E7A6ED; nxt/gateway.dll/vid=sefaz_tributaria%3Avtribut");

                    WebResponse res = req.GetResponse();
                    Stream dataStream = res.GetResponseStream();
                    StreamReader reader = new StreamReader(dataStream);

                    string resposta = reader.ReadToEnd().Replace("\"", string.Empty).Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");

                    var listaUrls = Regex.Split(resposta, "href=").ToList();

                    listaUrls.RemoveRange(0, 103);
                    listaUrls.RemoveAll(x => !x.Substring(0, x.IndexOf(">")).ToLower().Contains("info.fazenda.sp.gov.br/") || x.Contains("http://info.fazenda.sp.gov.br/NXT/gateway.dll/legislacao_tributaria/agendas/legistrib2.htm"));

                    listaUrls.ForEach(delegate(string itemDado)
                    {
                        try
                        {
                            objUrll = new ExpandoObject();

                            itemDado = itemDado.Substring(0, itemDado.IndexOf(">")).Replace(" target=_blank", string.Empty);

                            objUrll.Indexacao = "Secretaria do estado da fazenda de SP";
                            objUrll.Url = itemDado;
                            objUrll.Lista_Nivel2 = new List<dynamic>();

                            req = WebRequest.Create(itemDado);
                            req.Headers.Set("Cookie", "nxtopts=420zhlf0zzzzz; nxt/gateway.dll/uid=36E76821; nxt/gateway.dll/vid=sefaz_tributaria%3Avtribut");

                            res = req.GetResponse();
                            dataStream = res.GetResponseStream();
                            reader = new StreamReader(dataStream, Encoding.Default);

                            resposta = reader.ReadToEnd().Replace("\"", string.Empty).Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");
                            resposta = removeTagScript(resposta);

                            var listFinalURl = Regex.Split(resposta, "href=").ToList();

                            listFinalURl.RemoveRange(0, 2);
                            listFinalURl.RemoveAll(x => !x.Substring(0, x.IndexOf(">")).ToLower().Contains("info.fazenda.sp.gov.br/") && x.Substring(0, x.IndexOf(">")).Contains("/"));

                            listFinalURl.ForEach(delegate(string itemFinal)
                            {
                                objItenUrl = new ExpandoObject();

                                itemFinal = itemFinal.Substring(0, itemFinal.IndexOf(">"));

                                if (!itemFinal.Contains("/"))
                                    itemFinal = itemDado.Substring(0, itemDado.LastIndexOf("/") + 1) + itemFinal;

                                objItenUrl.Url = itemFinal;
                                objUrll.Lista_Nivel2.Add(objItenUrl);
                            });

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                        }
                        catch (Exception)
                        {
                        }
                    });
                }
            }

            #endregion

            #region "Captura DOC's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("d") && siglaFonteProcessamento.Equals("sefazSp"))
            {
                this.Text = "Fonte: Sefaz SP";

                string hash = string.Empty;
                string texto = string.Empty;
                string titulo = string.Empty;
                string publicacao = string.Empty;
                string ementa = string.Empty;
                string urlTratada = string.Empty;

                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazSp");

                var listaExt = new List<string>();
                dynamic ementaInserir;

                this.label1.Text = "Status: Processando";

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            hash = string.Empty;
                            texto = string.Empty;
                            titulo = string.Empty;
                            publicacao = string.Empty;
                            ementa = string.Empty;

                            ementaInserir = new ExpandoObject();
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            WebRequest req = WebRequest.Create(urlTratada);

                            if (!urlTratada.Contains("sp.gov.br/NXT/gateway.dll/Respostas_CT"))
                                req.Headers.Set("Cookie", "nxtopts=420zhlf0zzzzz; nxt/gateway.dll/uid=36E76821; nxt/gateway.dll/sid=36E7A6ED; nxt/gateway.dll/vid=sefaz_tributaria%3Avtribut");
                            else
                            {
                                hash = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1);
                                req.Headers.Set("Cookie", "nxtopts=420zhlf0zzzzz; nxt/gateway.dll/uid=36E97A5F; nxt/gateway.dll/sid=36E97ACE; nxt/gateway.dll/vid=sefaz_respct%3Avrespct; nxt/gateway.dll/doc=Respostas_CT%2Ficms%2F" + hash);
                            }

                            WebResponse res = req.GetResponse();
                            Stream dataStream = res.GetResponseStream();
                            StreamReader reader = new StreamReader(dataStream, Encoding.Default);

                            string respostaNivel2_ = reader.ReadToEnd().Replace("\"", string.Empty).Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");

                            List<string> listString = new List<string>();

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", string.Empty).Replace("\n", " ").Replace("\r", " ").Replace("\t", " ")).ToLower();

                            listString.Add(respostaNivel2_.IndexOf("<body") + "|<body");
                            listString.Add(respostaNivel2_.IndexOf("</head>") + "|</head>");

                            listString.RemoveAll(x => x.Contains("-1"));

                            texto = respostaNivel2_.Substring(respostaNivel2_.IndexOf(listString[0].Split('|')[1]));

                            if (respostaNivel2_.Contains("<title"))
                            {
                                titulo = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<title")).Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<title")).IndexOf("</title>"));
                                titulo = ObterStringLimpa(titulo);

                                listString = new List<string>();

                                //listString.Add(respostaNivel2_.IndexOf("</p>") + "|</p>");
                                //listString.Add(respostaNivel2_.IndexOf("<p class=") + "|<p class=");

                                //listString.RemoveAll(x => x.Contains("-1"));
                                //listString = listString.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                                if (texto.ToLower().Contains("doe "))
                                    publicacao = texto.Substring(texto.IndexOf("doe ")).Substring(0, texto.Substring(texto.IndexOf("doe ")).IndexOf("<"));
                                else if (titulo.Contains(","))
                                    publicacao = titulo.Substring(titulo.IndexOf(",") + 1);

                                publicacao = ObterStringLimpa(publicacao);
                            }
                            else
                            {
                                titulo = string.Empty;
                                publicacao = string.Empty;
                            }

                            hash = GerarHash(removerCaracterEspecial(removeTagScript(Regex.Replace(texto, @"\s+", " "))));

                            /*Captura CSS*/

                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (respostaNivel2_.Contains("<style"))
                                descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                           .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                            /**Fim Captura CSS**/
                            texto = novaCss + descStyle + texto;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "SP") + "}";

                            string numero = string.Empty;
                            string especie = string.Empty;

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = Regex.Replace(removeTagScript(texto), @"\s+", " ");
                            ementaInserir.Hash = hash;
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "SP";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > sefazSP > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }

                label1.Text = "• Status: Concluído";
            }

            #endregion

            #endregion

            #region "Pref. SP"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                this.Text = "Fonte: Pref. São Paulo";
                string urlRaiz = "http://legislacao.prefeitura.sp.gov.br/wp-json/cadlem_api/v1/busca?tipo[]=decreto&tipo[]=lei&tipo[]=orientacao-normativa&tipo[]=portaria&tipo[]=portaria-conjunta&tipo[]=portaria-intersecretarial&tipo[]=razoes-de-veto-oficios-atl&rppag=30&npag={0}";

                label1.Text = "Processando...";
                label1.Dispose();

                int maxPages = 0;
                int totalPages = 0;

                objUrll = new ExpandoObject();

                objUrll.Indexacao = "Prefeitura do Município de São Paulo";
                objUrll.Url = urlRaiz;

                while (maxPages <= totalPages)
                {
                    try
                    {
                        maxPages++;

                        Thread.Sleep(2000);

                        WebRequest reqNivel2_1 = WebRequest.Create(string.Format(urlRaiz, maxPages.ToString()));

                        WebResponse resNivel2_1 = reqNivel2_1.GetResponse();
                        Stream dataStreamNivel2_1 = resNivel2_1.GetResponseStream();
                        StreamReader readerNivel2_1;

                        string headerKey = resNivel2_1.Headers.Get("X-WP-TotalPages").ToString();

                        if (totalPages == 0) int.TryParse(headerKey, out totalPages);

                        readerNivel2_1 = new StreamReader(dataStreamNivel2_1, Encoding.Default);

                        string respostaNivel2_1 = readerNivel2_1.ReadToEnd();
                        var objDados = JObject.Parse(respostaNivel2_1);

                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        foreach (var item in objDados["posts"])
                        {
                            objItenUrl = new ExpandoObject();

                            objItenUrl.Url = item["link"].ToString();
                            objUrll.Lista_Nivel2.Add(objItenUrl);
                        }

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                    catch (Exception)
                    {
                    }
                }

                this.label1.Text = "Concluído";
            }

            #endregion

            #region "Captura DOC's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("d") && siglaFonteProcessamento.Equals("prefSp"))
            {
                this.Text = "Fonte: Pref SP";

                string hash = string.Empty;
                string texto = string.Empty;
                string titulo = string.Empty;
                string publicacao = string.Empty;
                string ementa = string.Empty;
                string urlTratada = string.Empty;

                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("prefSp");

                var listaExt = new List<string>();
                dynamic ementaInserir;

                this.label1.Text = "Status: Processando";

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            hash = string.Empty;
                            texto = string.Empty;
                            titulo = string.Empty;
                            publicacao = string.Empty;
                            ementa = string.Empty;

                            ementaInserir = new ExpandoObject();
                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            dynamic itemListaVez = new ExpandoObject();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string respostaNivel2_ = getHtmlPaginaByGet(urlTratada, string.Empty);

                            List<string> listString = new List<string>();

                            respostaNivel2_ = System.Net.WebUtility.HtmlDecode(respostaNivel2_.Replace("\"", string.Empty).Replace("\n", " ").Replace("\r", " ").Replace("\t", " "));

                            listString.Add(respostaNivel2_.ToLower().IndexOf("<div class=col-md-9 col-md-push-0 bx-ementa") + "|<div class=col-md-9 col-md-push-0 bx-ementa");

                            listString.RemoveAll(x => x.Contains("-1"));

                            if (listString.Count > 0)
                                texto = respostaNivel2_.Substring(int.Parse(listString[0].Split('|')[0])).Substring(0, respostaNivel2_.Substring(int.Parse(listString[0].Split('|')[0])).IndexOf("<div class=bx-mensagem"));

                            listString = new List<string>();

                            listString.Add(respostaNivel2_.ToLower().IndexOf("<h4>") + "|<h4>|</h4>");
                            listString.Add(respostaNivel2_.ToLower().IndexOf("<title>") + "|<title>|</title>");

                            listString.RemoveAll(x => x.Contains("-1"));

                            if (listString.Count > 0)
                            {
                                titulo = respostaNivel2_.Substring(int.Parse(listString[0].Split('|')[0])).Substring(0, respostaNivel2_.Substring(int.Parse(listString[0].Split('|')[0])).IndexOf(listString[0].Split('|')[2])).Trim();
                                titulo = ObterStringLimpa(titulo);

                                publicacao = titulo.Substring(titulo.ToLower().IndexOf(" de") + 4);

                                //publicacao = ObterStringLimpa(publicacao);
                            }
                            else
                            {
                                titulo = string.Empty;
                                publicacao = string.Empty;
                            }

                            hash = GerarHash(removerCaracterEspecial(removeTagScript(Regex.Replace(texto, @"\s+", " "))));

                            /*Captura CSS*/

                            var listaCss = Regex.Split(respostaNivel2_, "<link").ToList();
                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (respostaNivel2_.Contains("<style"))
                                descStyle = respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style"))
                                                           .Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("<style")).IndexOf("</style>") + 8);

                            /**Fim Captura CSS**/
                            texto = novaCss + descStyle + texto;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "SP") + "}";

                            string numero = string.Empty;
                            string especie = string.Empty;

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = string.Empty;
                            ementaInserir.Texto = Regex.Replace(removeTagScript(texto), @"\s+", " ");
                            ementaInserir.Hash = hash;
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "SP";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > prefSp > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }

                label1.Text = "Status: Concluído";
            }

            #endregion

            #endregion

            #region "Pref. RJ"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                this.Text = "Fonte: Pref. RJ";
                string linksPdf = string.Empty;

                string htmlUrl = getHtmlPaginaByGet("http://www2.rio.rj.gov.br/smf/fcet/legislacao.asp", "default").Replace("\"", string.Empty);

                if (htmlUrl.ToLower().Contains("<table class=tablebranco"))
                {
                    linksPdf = htmlUrl.Substring(htmlUrl.ToLower().IndexOf("<table class=tablebranco")).Substring(0, htmlUrl.Substring(htmlUrl.ToLower().IndexOf("<table class=tablebranco")).IndexOf("</table>"));

                    var listUrls = Regex.Split(linksPdf, "href=").ToList();

                    listUrls.RemoveAt(0);

                    objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Prefeitura do Município do Rio Janeiro";
                    objUrll.Url = "http://www2.rio.rj.gov.br/smf/fcet/legislacao.asp";

                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    listUrls.ForEach(delegate(string itemX)
                    {
                        objItenUrl = new ExpandoObject();

                        objItenUrl.Url = "http://www2.rio.rj.gov.br/smf/fcet/" + itemX.Substring(0, itemX.IndexOf(">"));
                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    });

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                }
            }

            #endregion

            #region "Captura DOC's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("d") && siglaFonteProcessamento.Equals("prefRj"))
            {
                this.Text = "Fonte: Pref RJ";

                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("prefRj");

                string hash = string.Empty;
                string texto = string.Empty;
                string titulo = string.Empty;
                string publicacao = string.Empty;
                string ementa = string.Empty;
                string urlTratada = string.Empty;
                string conteudoPdf = string.Empty;

                dynamic ementaInserir;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        Thread.Sleep(2000);

                        hash = string.Empty;
                        texto = string.Empty;
                        titulo = string.Empty;
                        publicacao = string.Empty;
                        ementa = string.Empty;

                        ementaInserir = new ExpandoObject();
                        ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                        dynamic itemListaVez = new ExpandoObject();

                        itemListaVez.ListaEmenta = new List<dynamic>();
                        urlTratada = itemLista_Nivel2.Url.Trim();
                        itemListaVez.Url = urlTratada;
                        itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                        string nomeArq = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1);

                        nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                        string numero = string.Empty;
                        string especie = string.Empty;

                        switch (urlTratada.Substring(urlTratada.LastIndexOf("/") + 1).Substring(0, urlTratada.Substring(urlTratada.LastIndexOf("/") + 1).IndexOf(".")))
                        {
                            case "ltrib_at": especie = "Leis Tributárias"; titulo = "Leis Tributárias - Consolidação";
                                break;
                            case "ltrib_ch": especie = "Leis Tributárias"; titulo = "Leis Tributárias - Consolidação";
                                break;
                            case "dec_chp01": especie = "DECRETOS TRIBUTÁRIOS"; titulo = "DECRETOS TRIBUTÁRIOS - CONSOLIDAÇÃO HISTÓRICA - PARTE I";
                                break;
                            case "dec_chp02": especie = "DECRETOS TRIBUTÁRIOS"; titulo = "DECRETOS TRIBUTÁRIOS - CONSOLIDAÇÃO HISTÓRICA - PARTE II";
                                break;
                            case "dreg_iss": especie = "DECRETOS TRIBUTÁRIOS"; titulo = "DECRETOS TRIBUTÁRIOS - REGULAMENTO DO ISS";
                                break;
                            case "dreg_iptu": especie = "DECRETOS TRIBUTÁRIOS"; titulo = "DECRETOS TRIBUTÁRIOS - REGULAMENTO DO IPTU";
                                break;
                            case "dpatrib": especie = "DECRETOS TRIBUTÁRIOS"; titulo = "DECRETOS TRIBUTÁRIOS - PROCESSO ADMINISTRATIVO TRIBUTÁRIO";
                                break;
                            case "rtrib_chp01": especie = "RESOLUÇÕES TRIBUTÁRIAS"; titulo = "RESOLUÇÕES TRIBUTÁRIAS - CONSOLIDAÇÃO HISTÓRICA - PARTE I";
                                break;
                            case "rtrib_chp02": especie = "RESOLUÇÕES TRIBUTÁRIAS"; titulo = "RESOLUÇÕES TRIBUTÁRIAS - CONSOLIDAÇÃO HISTÓRICA - PARTE II";
                                break;
                            case "port_ch": especie = "PORTARIAS TRIBUTÁRIAS"; titulo = "PORTARIAS TRIBUTÁRIAS - CONSOLIDAÇÃO HISTÓRICA";
                                break;
                            case "in_ch": especie = "INSTRUÇÕES NORMATIVAS"; titulo = "INSTRUÇÕES NORMATIVAS - CONSOLIDAÇÃO HISTÓRICA";
                                break;
                            case "cltrib_anual": especie = "CONSOLIDAÇÃO DAS LEIS TRIBUTÁRIAS"; titulo = "CONSOLIDAÇÃO DAS LEIS TRIBUTÁRIAS - CONSOLIDAÇÃO ANUAL";
                                break;
                            default:
                                break;
                        }

                        using (WebClient webclient = new WebClient())
                        {
                            webclient.DownloadFile(urlTratada, nomeArq);
                        }

                        using (ZipArchive archive = ZipFile.OpenRead(nomeArq))
                        {
                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                entry.ExtractToFile(nomeArq.Replace(".zip", ".pdf"));

                                conteudoPdf = LeArquivo(nomeArq.Replace(".zip", ".pdf"));
                            }
                        }

                        File.Delete(nomeArq);
                        File.Delete(nomeArq.Replace(".zip", ".pdf"));

                        hash = GerarHash(Regex.Replace(conteudoPdf, @"\s+", " "));

                        /** Outros **/
                        ementaInserir.DescSigla = string.Empty;
                        ementaInserir.HasContent = false;

                        /** Arquivo **/
                        ementaInserir.Tipo = 3;
                        ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "RJ") + "}";

                        titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                        especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                        if (numero.Equals(string.Empty))
                            numero = "0";

                        ementaInserir.Publicacao = publicacao;
                        ementaInserir.Sigla = string.Empty;
                        ementaInserir.Republicacao = string.Empty;
                        ementaInserir.Ementa = ementa;
                        ementaInserir.TituloAto = titulo;
                        ementaInserir.Especie = especie;
                        ementaInserir.NumeroAto = Regex.Replace("0", "[^0-9]+", string.Empty);
                        ementaInserir.DataEdicao = string.Empty;
                        ementaInserir.Texto = conteudoPdf;
                        ementaInserir.Hash = hash;
                        ementaInserir.IdFila = itemLista_Nivel2.Id;

                        ementaInserir.Escopo = "RJ";
                        ementaInserir.IdFila = itemLista_Nivel2.Id;

                        itemListaVez.ListaEmenta.Add(ementaInserir);

                        dynamic itemFonte = new ExpandoObject();

                        itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                    }
                }
            }

            #endregion

            #endregion

            #region "Sec. Mun. SP"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                this.Text = "Fonte: Sec.Mun.SP";
                string linksPdf = string.Empty;

                List<string> listaUrlSecSp = new List<string>() { "3160", "3159", "3161", "3162", "3163", "3164", "3165", "3166", "3167", "3168", "3169", "3171", "3172", "3795", "6651", "7093" };

                foreach (var item in listaUrlSecSp)
                {
                    string htmlUrl = getHtmlPaginaByGet("http://www.prefeitura.sp.gov.br/cidade/secretarias/fazenda/legislacao/index.php?p=" + item, "default").Replace("\"", string.Empty);

                    if (htmlUrl.ToLower().Contains("<div id=texto"))
                    {
                        linksPdf = htmlUrl.Substring(htmlUrl.ToLower().IndexOf("<div id=texto"));

                        var listUrls = Regex.Split(linksPdf, "href=").ToList();

                        listUrls.RemoveAt(0);
                        listUrls.RemoveAll(x => !x.ToLower().Contains("prefeitura.sp.gov.br") || x.ToLower().Contains("voltar"));

                        objUrll = new ExpandoObject();

                        objUrll.Indexacao = "Secretaria Municipal da Fazenda - São Paulo";
                        objUrll.Url = "http://www.prefeitura.sp.gov.br/cidade/secretarias/fazenda/legislacao/index.php?p=" + item;

                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        //Pegando os itens que já são .PDF
                        listUrls.FindAll(x => x.Substring(0, x.IndexOf(">")).Contains(".pdf") || !x.Substring(0, x.IndexOf(">")).Contains("?p=")).ToList().ForEach(delegate(string itemX)
                        {
                            objItenUrl = new ExpandoObject();

                            objItenUrl.Url = itemX.Substring(0, itemX.IndexOf(">"));
                            objUrll.Lista_Nivel2.Add(objItenUrl);
                        });

                        listUrls.FindAll(x => !x.Substring(0, x.IndexOf(">")).Contains(".pdf") && x.Substring(0, x.IndexOf(">")).Contains("?p=")).ToList().ForEach(delegate(string itemX)
                        {
                            htmlUrl = getHtmlPaginaByGet(itemX, "default").Replace("\"", string.Empty);

                            if (htmlUrl.ToLower().Contains("<div id=texto"))
                            {
                                linksPdf = htmlUrl.Substring(htmlUrl.ToLower().IndexOf("<div id=texto"));

                                var listUrls1 = Regex.Split(linksPdf, "href=").ToList();

                                listUrls1.RemoveAt(0);
                                listUrls1.RemoveAll(x => !x.ToLower().Contains("prefeitura.sp.gov.br") || x.ToLower().Contains("voltar"));

                                listUrls1.ForEach(delegate(string itemXx)
                                {
                                    objItenUrl = new ExpandoObject();

                                    objItenUrl.Url = itemXx.Substring(0, itemXx.IndexOf(">"));
                                    objUrll.Lista_Nivel2.Add(objItenUrl);
                                });
                            }
                        });

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("secMunSp"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("secMunSp");

                string urlTratada = string.Empty;
                int count = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (count < 18)
                            {
                                count++;
                                continue;
                            }
                            else
                                Thread.Sleep(2000);

                            count++;

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string dadosPdf = string.Empty;
                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;

                            if (!urlTratada.Contains("?p="))
                            {
                                string nomeArq = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1);
                                //string nomeForRep = nomeArq.Substring(0, nomeArq.LastIndexOf("."));

                                nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                                using (WebClient webclient = new WebClient())
                                {
                                    webclient.DownloadFile(urlTratada, nomeArq);
                                }

                                dadosPdf = LeArquivo(nomeArq);

                                var listDados = Regex.Split(dadosPdf, "\n").ToList();

                                listDados.RemoveAll(x => x.Trim().Equals(string.Empty) || Char.IsNumber(Convert.ToChar(x.Trim().Substring(0, 1))));

                                titulo = listDados[0];
                                publicacao = titulo.IndexOf("(") > 0 ? titulo.Substring(titulo.IndexOf("(") + 1) : listDados.Find(x => x.Contains("(DOC"));

                                titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });
                                especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                                File.Delete(nomeArq);
                            }
                            else
                            {
                                dadosPdf = getHtmlPaginaByGet(urlTratada, "default").Replace("\r", " ").Replace("\t", " ").Replace("\n", " ").Replace("\"", string.Empty);

                                string dadosCorpo = dadosPdf.Substring(dadosPdf.IndexOf("<div id=texto"));
                                titulo = ObterStringLimpa(dadosCorpo.Substring(dadosCorpo.IndexOf("<h2")).Substring(0, dadosCorpo.Substring(dadosCorpo.IndexOf("<h2")).IndexOf("</h2>")));
                                especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();
                            }

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "SP") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = dadosPdf;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(dadosPdf));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "SP";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > secMunSp > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz Acre"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                this.Text = "Fonte: Sefaz AC";

                linksSefazAc = new List<string>() { "http://www.sefaz.ac.gov.br/wps/portal/sefaz/sefaz/principal/!ut/p/c5/7ZLLboJQEIafxRfgDByBwxLSgxxuilxENgSIJYCAiY1Qnr5oF00XtYsaV53ZTDLzzzfJ_ChBc3bZpSqzt6rvsiOKUSKlG4N5fmiIPAm3GNgmWr9YKxETD8_9vZTCD6HCL-odimGZ-vX7iU3NtK1htNxcZ24eDg5teCcoRlf2waXnMaDN5NSa5YTldJ5UgELhI92jqt294oO6-OMl_-rnq02UlMc-n122u_ruziz_2b9Dco2-PaA9SuSvLWuHEmAippGuefxcoeCBfvvOIrYszAQG2DFBEFby81hk-VDW_JUqb7mhaDngiIyxuBQUUCRFkgSCYg2d2vBi3cIWTWPgyTUHdfEBGlrzCA!!/dl3/d3/L2dBISEvZ0FBIS9nQSEh/?1dmy&current=true&urile=wcm%3apath%3a/SefazServ/Portal+SefazServ/Principal/Servicos/Legislacao/Legislacao+Estadual+repo/Leis+Ordinarias/"
                                                   ,"http://www.sefaz.ac.gov.br/wps/portal/sefaz/sefaz/principal/!ut/p/c5/7ZLPboJAEMafxRdgd1kWliO2q_xxwVUWkQtB0xBQxKSNxH36LunB9FCbtKanzlwmmfnmN8l8oAA6T9Wlqau3pj9VR5CDwi6XfiDW0ieIyhWGwTJLnqM5wVRg3d_aJfwiPPiNegNyaJXr9noO1EGtWiWG1OEwVnzg7IB4yq68rTBnJEmlpetpxGWtXpUH4d5F2UwwLxayKvvJLy_5V_9EHYKiPvY77ZPN6Jw7s-ijf4cU-333AragcG5bEs4oDAhm2WwqkK5A-kDHfGbRhWNqQgAxD6Fpzp2_Y1HroSz9lWbXGcO-M6BBHYyJZbrQtV3bNinIn8C5k5dojAUJ_QHRMQdvMnkHWY3K_w!!/dl3/d3/L2dBISEvZ0FBIS9nQSEh/?1dmy&current=true&urile=wcm%3apath%3a/SefazServ/Portal+SefazServ/Principal/Servicos/Legislacao/Legislacao+Estadual+repo/Leis+Complementares/"
                                                   ,"http://www.al.ac.leg.br/leis/?page_id=9789"
                                                   ,"http://www.sefaz.ac.gov.br/wps/portal/sefaz/sefaz/principal/!ut/p/c5/7ZLPboJAEMafxRdgd1kWliO2q_xxwVUWkQtB0xBQxKSNxH36LunB9FCbtKanzlwmmfnmN8l8oAA6T9Wlqau3pj9VR5CDwi6XfiDW0ieIyhWGwTJLnqM5wVRg3d_aJfwiPPiNegNyaJXr9noO1EGtWiWG1OEwVnzg7IB4yq68rTBnJEmlpetpxGWtXpUH4d5F2UwwLxayKvvJLy_5V_9EHYKiPvY77ZPN6Jw7s-ijf4cU-333AragcG5bEs4oDAhm2WwqkK5A-kDHfGbRhWNqQgAxD6Fpzp2_Y1HroSz9lWbXGcO-M6BBHYyJZbrQtV3bNinIn8C5k5dojAUJ_QHRMQdvMnkHWY3K_w!!/dl3/d3/L2dBISEvZ0FBIS9nQSEh/?1dmy&current=true&urile=wcm%3apath%3a/SefazServ/Portal+SefazServ/Principal/Servicos/Legislacao/Legislacao+Estadual+repo/Decretos/"
                                                   ,"http://www.sefaz.ac.gov.br/wps/portal/sefaz/sefaz/principal/!ut/p/c5/7ZLPboJAEMafxRdgd1kWliO2q_xxwVUWkQtB0xBQxKSNxH36LunB9FCbtKanzlwmmfnmN8l8oAA6T9Wlqau3pj9VR5CDwi6XfiDW0ieIyhWGwTJLnqM5wVRg3d_aJfwiPPiNegNyaJXr9noO1EGtWiWG1OEwVnzg7IB4yq68rTBnJEmlpetpxGWtXpUH4d5F2UwwLxayKvvJLy_5V_9EHYKiPvY77ZPN6Jw7s-ijf4cU-333AragcG5bEs4oDAhm2WwqkK5A-kDHfGbRhWNqQgAxD6Fpzp2_Y1HroSz9lWbXGcO-M6BBHYyJZbrQtV3bNinIn8C5k5dojAUJ_QHRMQdvMnkHWY3K_w!!/dl3/d3/L2dBISEvZ0FBIS9nQSEh/?1dmy&current=true&urile=wcm%3apath%3a/SefazServ/Portal+SefazServ/Principal/Servicos/Legislacao/Legislacao+Estadual+repo/Portarias/"
                                                   ,"http://www.sefaz.ac.gov.br/wps/portal/sefaz/sefaz/principal/!ut/p/c5/7ZLPboJAEMafxRdgd1kWliO2q_xxwVUWkQtB0xBQxKSNxH36LunB9FCbtKanzlwmmfnmN8l8oAA6T9Wlqau3pj9VR5CDwi6XfiDW0ieIyhWGwTJLnqM5wVRg3d_aJfwiPPiNegNyaJXr9noO1EGtWiWG1OEwVnzg7IB4yq68rTBnJEmlpetpxGWtXpUH4d5F2UwwLxayKvvJLy_5V_9EHYKiPvY77ZPN6Jw7s-ijf4cU-333AragcG5bEs4oDAhm2WwqkK5A-kDHfGbRhWNqQgAxD6Fpzp2_Y1HroSz9lWbXGcO-M6BBHYyJZbrQtV3bNinIn8C5k5dojAUJ_QHRMQdvMnkHWY3K_w!!/dl3/d3/L2dBISEvZ0FBIS9nQSEh/?1dmy&current=true&urile=wcm%3apath%3a/SefazServ/Portal+SefazServ/Principal/Servicos/Legislacao/Legislacao+Estadual+repo/Instrucao+Normativa/"};

                webBrowser1 = new WebBrowser();
                webBrowser1.ScriptErrorsSuppressed = true;
                webBrowser1.Navigate(linksSefazAc[0]);
                webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted_SefazAc);
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazAc"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazAc");

                string urlTratada = string.Empty;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string dadosPdf = string.Empty;
                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;

                            string nomeArq = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1);
                            //string nomeForRep = nomeArq.Substring(0, nomeArq.LastIndexOf("."));

                            nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), Regex.Replace(nomeArq, "[^0-9a-zA-Z]+", string.Empty) + ".pdf");

                            using (WebClient webclient = new WebClient())
                            {
                                webclient.DownloadFile(urlTratada, nomeArq);
                            }

                            dadosPdf = LeArquivo(nomeArq);

                            var listDados = Regex.Split(dadosPdf, "\n").ToList();

                            listDados.RemoveAll(x => x.Trim().Equals(string.Empty) || x.Contains("ESTADO DO"));

                            titulo = listDados[0];
                            publicacao = listDados.Find(x => x.Contains("Publicad"));

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });
                            especie = titulo.Substring(0, obterPontoCorte(titulo)).Trim();

                            File.Delete(nomeArq);

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "AC") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = dadosPdf;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(dadosPdf));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "AC";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Equals("O servidor remoto retornou um erro: (404) Não Localizado."))
                                ex = ex;

                            ex.Source = "Docs > sefazAc > lv1";
                            new BuscaLegalDao().InserirLogErro(ex, urlTratada, string.Empty);
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz Amazonia"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                this.Text = "Fonte: Sefaz Amazonas";

                List<string> listLinksSefazAm = new List<string>(){
                                                     "http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Constitui%C3%A7%C3%A3o%20Estadual/Ano%201989/Arquivo/CE%201989.htm|f"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Lei%20Complementar%20Estadual/Lei%20Complementar%20Estadual.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Lei%20Estadual/Lei%20Estadual.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Decreto%20Estadual/Decreto%20Estadual.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Resolu%C3%A7%C3%A3o%20GSEFAZ/Resolu%C3%A7%C3%A3o%20GSEFAZ.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Portaria%20SER/Portaria%20GSEFAZ.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Aviso%20GSEFAZ/AVISO%20GSEFAZ.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Portaria%20GSEFAZ_GPGE/Portaria%20GSEFAZ-GPGE.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Portaria%20GSEFAZ_GSEPLANCTI/Portaria%20GSEFAZ-GSEPLANCTI.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Portaria%20GSEFAZ-GSEPLAN/Portaria%20GSEFAZ-GSEPLAN.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Portaria%20GSEFAZ_GSSP/Portaria%20GSEFAZ-GSSP.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Portaria%20GSEFAZ-GSSP-IPEM/Portaria%20GSEFAZ-GSSP-IPEM.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Portaria%20GSEPLANCTI/Portaria%20GSEPLANCTI.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Resolu%C3%A7%C3%A3o%20GSEFAZ-SIC/Resolu%C3%A7%C3%A3o%20GSEFAZ-SIC.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Resolu%C3%A7%C3%A3o%20GSEFAZ-SUSAM/Resolu%C3%A7%C3%A3o%20GSEFAZ-SUSAM.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Resolu%C3%A7%C3%A3o%20GSEFAZ-GSEPLAN/Resolu%C3%A7%C3%A3o%20GSEFAZ-GSEPLAN.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Resolu%C3%A7%C3%A3o%20GSEPLANCTI-GSEFAZ/RESOLU%C3%87%C3%83O%20GSEPLANCTI-GSEFAZ.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Resolu%C3%A7%C3%A3o%20GSEPLAN/Resolu%C3%A7%C3%A3o%20GSEPLAN.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Resolu%C3%A7%C3%A3o%20CODAM/Resolu%C3%A7%C3%A3o%20CODAM.htm"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Decreto%20Estadual/Ano%201999/Arquivo/DE_20686_99.htm|f"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Decreto%20Estadual/Ano%202006/Arquivo/DE_26428_06.htm|f"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Decreto%20Estadual/Ano%201979/Arquivo/DE_4564_79.htm|f"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Lei%20Estadual/Ano%202003/Arquivo/LE_2826_03.htm|f"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Decreto%20Estadual/Ano%202003/Arquivo/DE_23994_03.htm|f"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Lei%20Estadual/Ano%202007/Arquivo/LE%203151%2007.htm|f"
                                                    ,"http://online.sefaz.am.gov.br/silt/Normas/Legisla%C3%A7%C3%A3o%20Estadual/Decreto%20Estadual/Ano%201991/Arquivo/DE_14168_91.htm|f"};

                foreach (var item in listLinksSefazAm)
                {
                    try
                    {
                        objUrll = new ExpandoObject();

                        objUrll.Indexacao = "Secretaria do estado da Fazendo do Amazonas";
                        objUrll.Url = item;

                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        if (item.Contains("|f"))
                        {
                            objItenUrl = new ExpandoObject();

                            objItenUrl.Url = item;
                            objUrll.Lista_Nivel2.Add(objItenUrl);
                        }
                        else
                        {
                            string textHtml = getHtmlPaginaByGet(item, "default").Replace("\"", string.Empty).Replace("\t", " ").Replace("\r", " ");
                            textHtml = textHtml.Substring(textHtml.IndexOf("<div align=center>")).Substring(0, textHtml.Substring(textHtml.IndexOf("<div align=center>")).IndexOf("</table>"));

                            var listLinks = Regex.Split(textHtml.ToLower(), "href=").ToList();
                            listLinks.RemoveAt(0);

                            listLinks.ForEach(delegate(string xItem)
                            {
                                string novaUrl = string.Empty;

                                if (xItem.IndexOf(" ") < xItem.IndexOf(">"))
                                    novaUrl = item.Substring(0, item.LastIndexOf("/") + 1) + xItem.Substring(0, xItem.IndexOf(" "));
                                else
                                    novaUrl = item.Substring(0, item.LastIndexOf("/") + 1) + xItem.Substring(0, xItem.IndexOf(">"));

                                Thread.Sleep(2000);

                                textHtml = getHtmlPaginaByGet(novaUrl, "default").Replace("\"", string.Empty).Replace("\t", " ").Replace("\r", " ");
                                textHtml = textHtml.Substring(textHtml.IndexOf("<div align=center>")).Substring(0, textHtml.Substring(textHtml.IndexOf("<div align=center>")).IndexOf("</table>"));

                                var listLinks1 = Regex.Split(textHtml.ToLower(), "href=").ToList();
                                listLinks1.RemoveAt(0);

                                listLinks1.ForEach(delegate(string y)
                                {
                                    objItenUrl = new ExpandoObject();

                                    objItenUrl.Url = novaUrl.Substring(0, novaUrl.LastIndexOf("/") + 1) + y.Substring(0, y.IndexOf(" "));
                                    objUrll.Lista_Nivel2.Add(objItenUrl);
                                });
                            });
                        }

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazAm"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazAm");

                string urlTratada = string.Empty;

                int inferno = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (inferno < 1930)
                            {
                                inferno++;
                                continue;
                            }

                            inferno++;

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty);
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string htmlPage = getHtmlPaginaByGet(urlTratada, "default").Replace("\"", string.Empty).Replace("\t", " ").Replace("\r", " ").Replace("\n", " ");

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;

                            htmlPage = Regex.Replace(htmlPage.Trim(), @"\s+", " ");

                            if (htmlPage.Contains("<title>"))
                                titulo = ObterStringLimpa(htmlPage.Substring(htmlPage.IndexOf("<title>")).Substring(0, htmlPage.Substring(htmlPage.IndexOf("<title>")).IndexOf("</title>")));

                            if (titulo.Contains("LC") || urlTratada.ToLower().Contains("lei%20complementar") || urlTratada.ToLower().Contains("lei%20complementar"))
                                titulo = ObterStringLimpa(htmlPage.Substring(htmlPage.ToLower().IndexOf("lei complementar n")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("lei complementar n")).IndexOf("</p>")));

                            else if (titulo.ToLower().Contains("resolução") || urlTratada.ToLower().Contains("resolução%20codam") || urlTratada.ToLower().Contains("resolução%20gsefaz") || urlTratada.ToLower().Contains("resolu%c3%a7%c3%a3o%20gsefaz"))
                            {
                                if (htmlPage.ToLower().IndexOf("resolução n") > 0)
                                    titulo = ObterStringLimpa(htmlPage.Substring(htmlPage.ToLower().IndexOf("resolução n")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("resolução n")).IndexOf("</p>")));
                                else if (htmlPage.ToLower().IndexOf("resolução <span") > 0)
                                    titulo = ObterStringLimpa(htmlPage.Substring(htmlPage.ToLower().IndexOf("resolução <span")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("resolução <span")).IndexOf("</p>")));
                                else if (htmlPage.ToLower().IndexOf("resolução <o:p>") > 0)
                                    titulo = ObterStringLimpa(htmlPage.Substring(htmlPage.ToLower().IndexOf("resolução <o:p>")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("resolução <o:p>")).IndexOf("</p>")));
                            }

                            else if (titulo.ToLower().Contains("decreto") || urlTratada.ToLower().Contains("decreto%20estadual"))
                                titulo = htmlPage.ToLower().IndexOf("decreto n") > 0 ? ObterStringLimpa(htmlPage.Substring(htmlPage.ToLower().IndexOf("decreto n")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("decreto n")).IndexOf("</p>"))) : titulo;

                            else if (titulo.ToLower().Contains("portaria") || urlTratada.ToLower().Contains("portaria%20ser"))
                                titulo = htmlPage.ToLower().IndexOf("portaria<o:p>") > 0 ? ObterStringLimpa(htmlPage.Substring(htmlPage.ToLower().IndexOf("portaria<o:p>")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("portaria<o:p>")).IndexOf("</p>"))) : titulo;

                            else if (titulo.ToLower().Contains("lei") || urlTratada.ToLower().Contains("lei%20estadual"))
                                titulo = htmlPage.ToLower().IndexOf("lei n") > 0 ? ObterStringLimpa(htmlPage.Substring(htmlPage.ToLower().IndexOf("lei n")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("lei n")).IndexOf("</p>"))) : titulo;

                            else
                                titulo = titulo;

                            List<int> listPublicacao = new List<int>() { titulo.ToLower().IndexOf(","), titulo.ToLower().IndexOf(" de ") };

                            listPublicacao.RemoveAll(x => x == -1);

                            if (listPublicacao.Count > 0)
                                publicacao = titulo.Substring(listPublicacao.Min(x => x) + 4);

                            else if (htmlPage.ToLower().Contains("publicad"))
                                publicacao = ObterStringLimpa(htmlPage.Substring(htmlPage.ToLower().IndexOf("publicad")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("publicad")).IndexOf("</p>")));

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });
                            texto = htmlPage.Substring(htmlPage.ToLower().IndexOf("<body")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("<body")).IndexOf("</body>"));

                            /*Captura CSS*/

                            var listaCss = Regex.Split(htmlPage, "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (htmlPage.Contains("<style"))
                                descStyle = htmlPage.Substring(htmlPage.IndexOf("<style"))
                                                           .Substring(0, htmlPage.Substring(htmlPage.IndexOf("<style")).IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            texto = novaCss + descStyle + texto;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "AM") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(removeTagScript(texto))));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "AM";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz Amapa"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                this.Text = "Fonte: Sefaz Amazonas";

                List<string> listLinks = new List<string>() { "https://www.sefaz.ap.gov.br/index.php/regulamentos/7894-ricms-decreto-n-2269-de-24-de-julho-de-1998"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/regulamentos/7895-ipva-decreto-n-3340-de-14-de-dezembro-de-1995"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/regulamentos/7896-itcd-decreto-n-3601-de-29-de-dezembro-de-2000"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/regulamentos/7897-taxas-decreto-n-2269-de-29-de-dezembro-de-1993"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/regulamentos/7901-taxas-decreto-n-7907-de-29-de-dezembro-de-2003"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/leis"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/outros-atos-legais/decretos"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/outros-atos-legais/portarias"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/outros-atos-legais/i-normativas"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/outros-atos-legais/atos-declaratorios"
                                                             ,"https://www.sefaz.ap.gov.br/index.php/constituicao-estadual﻿" };

                foreach (string itemLista in listLinks)
                {
                    string corpoLinks = string.Empty;

                    objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Secretaria do estado da Fazenda do Amapa";
                    objUrll.Url = itemLista;

                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    if (Char.IsNumber(Convert.ToChar(itemLista.Substring(itemLista.Length - 1))))
                    {
                        objItenUrl = new ExpandoObject();

                        objItenUrl.Url = itemLista;
                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    }
                    else
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            string htmlCode = getHtmlPaginaByGet(itemLista, string.Empty).Replace("\"", string.Empty).Replace("\n", " ").Replace("\r", " ").Replace("\t", " ");

                            List<string> listQuebra = new List<string>() { htmlCode.ToLower().IndexOf("<div id=ariext123_container").ToString() +  "|<div id=ariext123_container"
                                                                          ,htmlCode.ToLower().IndexOf("<div id=ariext124_container").ToString() +  "|<div id=ariext124_container"
                                                                          ,htmlCode.ToLower().IndexOf("<div id=ariext125_container").ToString() +  "|<div id=ariext125_container"
                                                                          ,htmlCode.ToLower().IndexOf("<div id=ariext126_container").ToString() +  "|<div id=ariext126_container"
                                                                          ,htmlCode.ToLower().IndexOf("<div id=ariext127_container").ToString() +  "|<div id=ariext127_container"
                                                                          ,htmlCode.ToLower().IndexOf("<dl class=article-info").ToString() +  "|<dl class=article-info"};

                            listQuebra.RemoveAll(x => x.Contains("-1"));

                            if (listQuebra.Count > 0)
                            {
                                Thread.Sleep(2000);

                                corpoLinks = htmlCode.Substring(int.Parse(listQuebra[0].Split('|')[0])).Substring(0, htmlCode.Substring(int.Parse(listQuebra[0].Split('|')[0])).IndexOf("</ul>"));

                                var listUrl = Regex.Split(corpoLinks, "href=").ToList();

                                listUrl.RemoveAt(0);

                                foreach (string itemUrl in listUrl)
                                {
                                    try
                                    {
                                        Thread.Sleep(2000);

                                        htmlCode = itemLista.Substring(0, itemLista.LastIndexOf("/index.php")) + itemUrl.Substring(0, itemUrl.IndexOf(" "));

                                        htmlCode = getHtmlPaginaByGet(htmlCode, string.Empty).Replace("\"", string.Empty).Replace("\n", " ").Replace("\r", " ").Replace("\t", " ");

                                        htmlCode = htmlCode.Substring(htmlCode.IndexOf("<div class=category-list")).Substring(0, htmlCode.Substring(htmlCode.IndexOf("<div class=category-list")).IndexOf("</table>"));

                                        if (htmlCode.ToLower().Contains("não há artigos nesta categoria"))
                                            continue;

                                        else
                                        {
                                            var listfinal = Regex.Split(htmlCode, "href=").ToList();

                                            listfinal.RemoveAt(0);
                                            listfinal.RemoveAll(x => !x.Substring(0, 1).Equals("/"));

                                            foreach (var item in listfinal)
                                            {
                                                if (!ObterStringLimpa("<" + item).ToLower().Contains("nº"))
                                                {
                                                    Thread.Sleep(2000);

                                                    htmlCode = itemLista.Substring(0, itemLista.LastIndexOf("/index.php")) + item.Substring(0, item.IndexOf(">"));
                                                    htmlCode = getHtmlPaginaByGet(htmlCode, string.Empty).Replace("\"", string.Empty).Replace("\n", " ").Replace("\r", " ").Replace("\t", " ");

                                                    htmlCode = htmlCode.Substring(htmlCode.ToLower().IndexOf("<dl class=article-info")).Substring(0, htmlCode.Substring(htmlCode.ToLower().IndexOf("<dl class=article-info")).IndexOf("</table>"));

                                                    var listfinal1 = Regex.Split(htmlCode, "href=").ToList();

                                                    listfinal1.RemoveAt(0);
                                                    listfinal1.RemoveAll(x => !x.Substring(0, 1).Equals("/") || !x.Substring(0, x.IndexOf(" ")).Contains(".pdf"));

                                                    if (listfinal1.Count > 0)
                                                    {
                                                        foreach (var item1 in listfinal1)
                                                        {
                                                            objItenUrl = new ExpandoObject();
                                                            objItenUrl.Url = itemLista.Substring(0, itemLista.LastIndexOf("/index.php")) + item1.Substring(0, item1.IndexOf(" "));
                                                            objUrll.Lista_Nivel2.Add(objItenUrl);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        objItenUrl = new ExpandoObject();
                                                        objItenUrl.Url = itemLista.Substring(0, itemLista.LastIndexOf("/index.php")) + item.Substring(0, item.IndexOf(">"));
                                                        objUrll.Lista_Nivel2.Add(objItenUrl);
                                                    }

                                                }
                                                else
                                                {
                                                    objItenUrl = new ExpandoObject();
                                                    objItenUrl.Url = itemLista.Substring(0, itemLista.LastIndexOf("/index.php")) + item.Substring(0, item.IndexOf(">"));
                                                    objUrll.Lista_Nivel2.Add(objItenUrl);
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }
                            }
                            else
                                break;
                        }
                        catch (Exception)
                        {
                        }
                    }

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazAp"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazAp");

                string urlTratada = string.Empty;

                int inferno = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (inferno < 828)
                            //if (!itemLista_Nivel2.Url.Trim().Equals("https://www.sefaz.ap.gov.br/index.php/ano1998/6404-decreto-n-0877-07-de-abril-de-1998-dispoe-sobre-normas-regulamentadoras-para-o-uso-obrigatorio-de-equipamento-emissor-de-cupom-fiscal-ecf-por-contribuintes-de-icms"))
                            {
                                inferno++;
                                continue;
                            }

                            inferno++;

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty);
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string dadosPdf = string.Empty;
                            string nomeArq = string.Empty;
                            string htmlPage = string.Empty;
                            string corpoDados = string.Empty;

                            if (urlTratada.Contains(".pdf"))
                            {
                                nomeArq = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1);

                                nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                                using (WebClient webClient = new WebClient())
                                {
                                    webClient.DownloadFile(urlTratada, nomeArq);
                                }

                                dadosPdf = LeArquivo(nomeArq);

                                especie = "Sefaz Amapá";
                                titulo = "Sefaz Amapá";

                                File.Delete(nomeArq);
                            }
                            else
                            {
                                htmlPage = getHtmlPaginaByGet(urlTratada, string.Empty).Replace("\"", string.Empty).Replace("\t", " ").Replace("\r", " ").Replace("\n", " ");
                                htmlPage = Regex.Replace(htmlPage.Trim(), @"\s+", " ");

                                corpoDados = htmlPage.Substring(htmlPage.IndexOf("<dl class=article-info")).Substring(0, htmlPage.Substring(htmlPage.IndexOf("<dl class=article-info")).IndexOf("<div class=userbottom"));

                                List<string> indexEmenta = new List<string>() { corpoDados.IndexOf("<p style=text-align: justify").ToString() + "|</div>"
                                                                               ,corpoDados.IndexOf("<p style=margin-right: 7.1pt;>").ToString() + "|</span>"
                                                                               ,corpoDados.IndexOf("<span style=font-family: 'Arial','sans-serif'; color: windowtext; font-size: 10pt;").ToString() + "|</span>"
                                                                               ,corpoDados.IndexOf("<td style=width: 52%; padding: 0cm; width=52%>").ToString() + "|</td>"
                                                                               ,corpoDados.IndexOf("<td width=52%>").ToString() + "|</td>"};

                                indexEmenta.RemoveAll(x => x.Contains("-1"));

                                titulo = htmlPage.Substring(htmlPage.ToLower().IndexOf("<title>")).Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("<title>")).IndexOf("</title>"));
                                titulo = ObterStringLimpa(titulo).Substring(0, ObterStringLimpa(titulo).IndexOf("-"));

                                if (string.IsNullOrEmpty(titulo))
                                    titulo = corpoDados.Substring(corpoDados.IndexOf("<strong>")).Substring(0, corpoDados.Substring(corpoDados.IndexOf("<strong>")).IndexOf("</strong>"));

                                titulo = ObterStringLimpa(titulo);

                                //titulo = corpoDados.Substring(corpoDados.IndexOf("<strong>")).Substring(0, corpoDados.Substring(corpoDados.IndexOf("<strong>")).IndexOf("</strong>"));
                                if (indexEmenta.Count > 0)
                                    ementa = corpoDados.Substring(int.Parse(indexEmenta[0].Split('|')[0])).Substring(0, corpoDados.Substring(int.Parse(indexEmenta[0].Split('|')[0])).IndexOf(indexEmenta[0].Split('|')[1]));

                                titulo = ObterStringLimpa(titulo);
                                ementa = ObterStringLimpa(ementa);

                                if (corpoDados.Contains(".doc") || corpoDados.Contains(".pdf"))
                                {
                                    string corteLink = urlTratada.Substring(0, urlTratada.IndexOf(".br/") + 3) + corpoDados.Substring(corpoDados.IndexOf("href=") + 5);

                                    indexEmenta = new List<string>() { corteLink.IndexOf(">").ToString(), corteLink.IndexOf(" ").ToString() };
                                    indexEmenta.RemoveAll(x => x.Contains("-1"));
                                    indexEmenta = indexEmenta.OrderBy(x => x).ToList();

                                    string linkArq = corteLink.Substring(0, int.Parse(indexEmenta[0]));

                                    nomeArq = linkArq.Substring(linkArq.LastIndexOf("/") + 1);
                                    nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                                    using (WebClient webClient = new WebClient())
                                    {
                                        webClient.DownloadFile(linkArq, nomeArq);
                                    }

                                    if (linkArq.Contains(".pdf"))
                                    {
                                        dadosPdf = LeArquivo(nomeArq);
                                    }
                                    else
                                    {
                                        byte[] arrayFile = File.ReadAllBytes(nomeArq);
                                        ementaInserir.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = nomeArq.Substring(nomeArq.LastIndexOf(".") + 1), NomeArquivo = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1) });
                                    }

                                    File.Delete(nomeArq);
                                }
                            }

                            List<int> listPublicacao = new List<int>() { titulo.ToLower().IndexOf(","), titulo.ToLower().IndexOf(" de ") };

                            listPublicacao.RemoveAll(x => x == -1);

                            if (listPublicacao.Count > 0)
                                publicacao = titulo.Substring(listPublicacao.Min(x => x) + 4);

                            titulo.ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            /*Captura CSS*/

                            var listaCss = Regex.Split(htmlPage, "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (htmlPage.Contains("<style"))
                                descStyle = htmlPage.Substring(htmlPage.IndexOf("<style"))
                                                           .Substring(0, htmlPage.Substring(htmlPage.IndexOf("<style")).IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            texto = novaCss + descStyle + corpoDados + dadosPdf;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "AP") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = string.Empty;
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(removeTagScript(texto))));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "AP";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz Goias"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                this.Text = "Fonte: Sefaz Goias";

                List<string> listLinksSefazAm = new List<string>() { 
                  /*"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Rcte/RCTE.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Leis_Diversas.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Decretos_Diversos.htm"  
                 ,*/"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Parecer_Normativo_SRE.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Instrucao_de_Servico_DIEF.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Portaria_GIEF.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Instrucao_de_Servico_DEAR.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Instrucao_de_Servico_GERC.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Instrucao_de_Servico_Conjunta_COF.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Instrucao_Normativa_GSF.htm"    
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Instrucao_Normativa_Conjuntas.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/indice_Ato_Normativo.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/Indice_Portaria_GSF.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Indices/indice_nota_oficial.htm"  
                 ,"ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/Secretario/MEMORANDO/M_064_2007.htm﻿"};

                foreach (var item in listLinksSefazAm)
                {
                    try
                    {
                        objUrll = new ExpandoObject();

                        objUrll.Indexacao = "Secretaria do estado da Fazendo do Goias";
                        objUrll.Url = item;

                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        string textHtml = getHtmlPaginaByGet(item, "default").Replace("\"", string.Empty).Replace("\t", " ").Replace("\r", " ");

                        var listLinks = Regex.Split(textHtml.ToLower(), "href=").ToList();
                        listLinks.RemoveAt(0);

                        listLinks.RemoveAll(x => !x.Substring(0, 4).ToLower().Contains("http") && !x.Substring(0, 4).ToLower().Contains("../") && !x.Substring(0, 4).ToLower().Contains(@"..\"));

                        listLinks.ForEach(delegate(string xItem)
                        {
                            string novaUrl = string.Empty;

                            if (!xItem.Substring(0, 4).ToLower().Contains("http"))
                                novaUrl = (item.Substring(0, item.IndexOf("/legislacao") + 11) + xItem.Substring(0, xItem.IndexOf(">"))).Replace("../", "/").Replace("..\\", "/").Replace(@"\", "/");
                            else
                                novaUrl = xItem;
                            objItenUrl = new ExpandoObject();

                            objItenUrl.Url = novaUrl;
                            objUrll.Lista_Nivel2.Add(objItenUrl);
                        });

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazGo"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazGo");

                string urlTratada = string.Empty;
                string htmlPage = string.Empty;

                int inferno = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (inferno < 45)
                            //if ("ftp://ftp.sefaz.go.gov.br/sefazgo/legislacao/superintendencia/sgaf/is/is_002_2006.htm" != itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty))
                            {
                                inferno++;
                                continue;
                            }

                            inferno++;

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty);
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string dadosPdf = string.Empty;
                            string nomeArq = string.Empty;
                            //string htmlPage = string.Empty;
                            string corpoDados = string.Empty;

                            htmlPage = getHtmlPaginaByGet(urlTratada, "default").Replace("\"", string.Empty).Replace("\t", " ").Replace("\r", " ").Replace("\n", " ").Replace("'", string.Empty);
                            htmlPage = Regex.Replace(htmlPage.Trim(), @"\s+", " ");

                            List<string> listMarcacao = new List<string>() {
                                htmlPage.ToLower().IndexOf("<p class=decretottulo").ToString() + "|</body>|<p class=decretottulo|<p class=doedata|<p class=parementa|</p>|<p class=doedata|</p>"
                               ,htmlPage.ToLower().IndexOf("<font face=arial color=#000080").ToString() + "|</body>|<font face=arial color=#000080|</font>|<font face=arial color=#800000|</font>|nd|nd"

                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=decretottulo") && htmlPage.ToLower().Contains("<p class=instrttulo") && (!htmlPage.ToLower().Contains("<p class=msonormal") || htmlPage.ToLower().Contains("<p class=esquerdo1") || htmlPage.ToLower().Contains("<p class=esquerdo2") || htmlPage.ToLower().Contains("<p class=notavalidade") || htmlPage.ToLower().Contains("<p class=esquerdo3")) ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=instrttulo|</p>|<p class=parementa|</p>|<span style=font-size:8.0pt;mso-bidi-font-size:10.0pt;|</span>"
                               ,(htmlPage.ToLower().IndexOf("<div class=wordsection1") >= 0 && !htmlPage.ToLower().Contains("<p class=decretottulo") && htmlPage.ToLower().Contains("<p class=instrttulo") ? htmlPage.ToLower().IndexOf("<div class=wordsection1").ToString() : "-1") + "|</body>|<p class=instrttulo|</p>|<p class=parementa|</p>|<span style=font-size:8.0pt;font-weight:normal;mso-bidi-font-weight:|</span>"
                               ,(htmlPage.ToLower().IndexOf("<div class=wordsection1") >= 0 && htmlPage.ToLower().Contains("<p class=textorevo") && !htmlPage.ToLower().Contains("<p class=parementa") ? htmlPage.ToLower().IndexOf("<div class=wordsection1").ToString() : "-1") + "|</body>|<p class=textorevo|</p>|nd|nd|(publicada|</p>"

                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=msobodytext") && !htmlPage.ToLower().Contains("<p class=titcentral") && !htmlPage.ToLower().Contains("<p class=link style=") && htmlPage.ToLower().Contains("<p class=doedata") && !htmlPage.ToLower().Contains("<p class=instrttulo") && !htmlPage.ToLower().Contains("<p class=ce ") && !htmlPage.ToLower().Contains("<p class=decretottulo") && !htmlPage.ToLower().Contains("<p class=leittulo") && !htmlPage.ToLower().Contains("<p class=mcitao") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=doedata|</p>|<p class=parementa|</p>|(publicada|</p>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=parnormal") && htmlPage.ToLower().Contains("<p class=instrttulo") && htmlPage.ToLower().Contains("<p class=msonormal") && !htmlPage.ToLower().Contains("<p class=ce ") && !htmlPage.ToLower().Contains("<p class=decretottulo") && !htmlPage.ToLower().Contains("<p class=leittulo") && !htmlPage.ToLower().Contains("<p class=mcitao") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=instrttulo|</p>|<p class=msonormal|</p>|(publicada|</span>"

                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=titcentral") && htmlPage.ToLower().Contains("<p class=msonormal") && htmlPage.ToLower().Contains("<p class=ce ") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=msonormal|</p>|<p class=parementa|</p>|(publicad|</span>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=leittulo") && htmlPage.ToLower().Contains("<p class=leittulo") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=leittulo|</p>|<p class=mementa|<p class=mfundamento|(d.o.e|</span>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=titcentral") && htmlPage.ToLower().Contains("<p class=mcitao") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=mcitao|</p>|<p class=mementa>|</p>|<p class=mementa align=center|</p>"
                                    
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=parnormal") && htmlPage.ToLower().Contains("<span style=font-size:12.0pt;mso-bidi-font-size:10.0pt;font-family:") && !htmlPage.ToLower().Contains("<p class=decretottulo") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=msonormal align=|</p>|<p class=msonormal style=margin-top|</p>|doe |</p>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && htmlPage.ToLower().Contains("<p class=textolivre") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=textolivre|</p>|<p class=ementa|</p>|doe |</p>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=titcentral") && !htmlPage.ToLower().Contains("<p class=leittulo") && !htmlPage.ToLower().Contains("<p class=link style=") && htmlPage.ToLower().Contains("<p class=msonormal align=center") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=msonormal align=center|</p>|<p class=msonormal style=margin-top:0cm;margin-right:0cm;margin-bottom:6.0pt;|<p class=MsoNormal style=margin-bottom:6.0pt;text-align:justify;text-indent: 4.0cm|doe |</p>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=msobodytext") && htmlPage.ToLower().Contains("<p class=msobodytext align=center") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=msobodytext align=center|</p>|<p class=msonormal style=margin-left:177.0pt;text-align:justify;tab-stops:|</p>|nd|nd"
                               ,(htmlPage.ToLower().IndexOf("<div class=wordsection1") >= 0 && !htmlPage.ToLower().Contains("<p class=decretottulo") && htmlPage.ToLower().Contains("<span style=mso-bookmark:") ? htmlPage.ToLower().IndexOf("<div class=wordsection1").ToString() : "-1") + "|</body>|<span style=mso-bookmark:|</span>|<p class=parementa|</p>|<p class=doedata|</p>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=msotitle") && !htmlPage.ToLower().Contains("<p class=decretottulo") && htmlPage.ToLower().Contains("<span style=font-family:arial; mso-bidi-font-family:times new roman") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<span style=font-family:arial; mso-bidi-font-family:times new roman|</span>|<p class=sm |</p>|(di|</span>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=parementa") && htmlPage.ToLower().Contains("<p class=parnormal") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=parnormal|</p>|<p class=titcentral|</p>|<p class=mepgrafe|</p>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && htmlPage.ToLower().Contains("<p class=titcentralrev") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=titcentralrev|</p>|</p> <p class=EmentaRev style=mso-pagination:widow-orphan>prorroga|</p> <p class=parrevoga style=mso-pagination:widow-orphan>|<p class=datadoe|</p>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && htmlPage.ToLower().Contains("<p class=link style=")  ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=link style=|</p>|<p class=parementa|</p>|<p class=doedata|</p>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && htmlPage.ToLower().Contains("<p class=leittulo")  ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=leittulo|</p>|<p class=parementa|</p>|<p class=doedata|</p>"
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && htmlPage.ToLower().Contains("<p class=msotitle") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=msotitle|</p>|<p class=parementa|</p>|<p class=doedata|</p>" 
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && htmlPage.ToLower().Contains("<p class=titcentral") && htmlPage.ToLower().Contains("<p class=mcitao") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=titcentral|</p>|<p class=parementa|</p>|<p class=doedata|</p>" 
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && !htmlPage.ToLower().Contains("<p class=ce ")  && htmlPage.ToLower().Contains("<p class=titcentral") && !htmlPage.ToLower().Contains("<p class=mcitao") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=titcentral|</p>|<p class=parementa|</p>|<p class=doedata|</p>" 
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && htmlPage.ToLower().Contains("<p class=ce ")  && htmlPage.ToLower().Contains("<p class=titcentral") && !htmlPage.ToLower().Contains("<p class=mcitao") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=ce |</p>|<p class=parementa|</p>|<p class=doedata|</p>" 
                               ,(htmlPage.ToLower().IndexOf("<div class=section1") >= 0 && htmlPage.ToLower().Contains("<p class=msobodytext") ? htmlPage.ToLower().IndexOf("<div class=section1").ToString() : "-1") + "|</body>|<p class=msobodytext|</p>|<p class=parementa|</p>|<p class=doedata|</p>" };

                            //,htmlPage.ToLower().IndexOf("").ToString() + "|||||||"

                            listMarcacao.RemoveAll(x => x.Contains("-1"));

                            listMarcacao = listMarcacao.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                            if (listMarcacao.Count != 1) listMarcacao = new List<string>();

                            if (htmlPage.Substring(int.Parse(listMarcacao[0].Split('|')[0])).ToLower().IndexOf(listMarcacao[0].Split('|')[1]) > 0)
                                corpoDados = htmlPage.Substring(int.Parse(listMarcacao[0].Split('|')[0]));
                            else
                                corpoDados = htmlPage.Substring(int.Parse(listMarcacao[0].Split('|')[0])).Substring(0, htmlPage.Substring(int.Parse(listMarcacao[0].Split('|')[0])).ToLower().IndexOf(listMarcacao[0].Split('|')[1]));

                            if (string.IsNullOrEmpty(titulo))
                                titulo = corpoDados.Substring(corpoDados.ToLower().IndexOf(listMarcacao[0].Split('|')[2])).Substring(0, corpoDados.Substring(corpoDados.ToLower().IndexOf(listMarcacao[0].Split('|')[2])).ToLower().IndexOf(listMarcacao[0].Split('|')[3]));

                            titulo = ObterStringLimpa(titulo);

                            if (listMarcacao.Count > 0 && !listMarcacao[0].Split('|')[4].Equals("nd") && corpoDados.ToLower().IndexOf(listMarcacao[0].Split('|')[4]) >= 0)
                                ementa = corpoDados.Substring(corpoDados.ToLower().IndexOf(listMarcacao[0].Split('|')[4])).Substring(0, corpoDados.Substring(corpoDados.ToLower().IndexOf(listMarcacao[0].Split('|')[4])).IndexOf(listMarcacao[0].Split('|')[5]));

                            titulo = ObterStringLimpa(titulo);
                            ementa = ObterStringLimpa(ementa);

                            if (corpoDados.ToLower().IndexOf(listMarcacao[0].Split('|')[6]) >= 0 && !listMarcacao[0].Split('|')[6].Equals("nd"))
                                publicacao = corpoDados.Substring(corpoDados.ToLower().IndexOf(listMarcacao[0].Split('|')[6])).Substring(0, corpoDados.Substring(corpoDados.ToLower().IndexOf(listMarcacao[0].Split('|')[6])).IndexOf(listMarcacao[0].Split('|')[7]));

                            if (publicacao.Equals(string.Empty))
                            {
                                List<int> listPublicacao = new List<int>() { titulo.ToLower().IndexOf(","), titulo.ToLower().IndexOf(" de ") };

                                listPublicacao.RemoveAll(x => x == -1);

                                if (listPublicacao.Count > 0)
                                    publicacao = titulo.Substring(listPublicacao.Min(x => x) + 4);
                            }

                            publicacao = ObterStringLimpa(publicacao);

                            titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            /*Captura CSS*/

                            var listaCss = Regex.Split(htmlPage, "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (htmlPage.Contains("<style"))
                                descStyle = htmlPage.Substring(htmlPage.IndexOf("<style"))
                                                           .Substring(0, htmlPage.Substring(htmlPage.IndexOf("<style")).IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            texto = novaCss + descStyle + corpoDados + dadosPdf;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "GO") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = "Sefaz GO";
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(removeTagScript(texto))));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "GO";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!htmlPage.Contains("Erro ao processar página, detalhes:"))
                                MessageBox.Show("Erro - index:" + inferno.ToString());
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz CE"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                this.Text = "Fonte: Sefaz CE";

                linksSefazCe = new List<string>() { "http://legis.sefaz.ce.gov.br/CGI-BIN/om_isapi.dll?clientID=5690479&hitsperheading=on&infobase=lei_complementar&record=toccurrent&softpage=TOC42"
                                                   ,"http://legis.sefaz.ce.gov.br/CGI-BIN/om_isapi.dll?clientID=5690523&hitsperheading=on&infobase=normas%20de%20execucao&record=toccurrent&softpage=TOC42"
                                                   ,"http://legis.sefaz.ce.gov.br/CGI-BIN/om_isapi.dll?clientID=5690524&hitsperheading=on&infobase=decretos&record=toccurrent&softpage=TOC42"
                                                   ,"http://legis.sefaz.ce.gov.br/CGI-BIN/om_isapi.dll?clientID=5690525&hitsperheading=on&infobase=instrucoes%20normativas&record=toccurrent&softpage=TOC42"
                                                   ,"http://legis.sefaz.ce.gov.br/CGI-BIN/om_isapi.dll?clientID=5690526&hitsperheading=on&infobase=leis&record=toccurrent&softpage=TOC42"
                                                   ,"http://legis.sefaz.ce.gov.br/CGI-BIN/om_isapi.dll?clientID=5690527&hitsperheading=on&infobase=notas%20explicativas&record=toccurrent&softpage=TOC42"
                                                   ,"http://legis.sefaz.ce.gov.br/CGI-BIN/om_isapi.dll?clientID=5690528&hitsperheading=on&infobase=decreto%2024569%20%28ricms%29&record=toccurrent&softpage=TOC42"};

                if (itemVez < linksSefazCe.Count)
                {
                    webBrowser1 = new WebBrowser();
                    webBrowser1.ScriptErrorsSuppressed = true;
                    webBrowser1.Navigate(linksSefazCe[itemVez]);
                    webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted_SefazCe);
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazCe"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazCe");

                string urlTratada = string.Empty;
                string htmlPage = string.Empty;

                int inferno = 0;

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (inferno < 0)
                            //if (!itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty).Equals("http://legis.sefaz.ce.gov.br/cgi-bin/om_isapi.dll?clientid=5690594&hitsperheading=on&infobase=decretos&record={17e5}&softpage=document42"))
                            {
                                inferno++;
                                continue;
                            }

                            inferno++;

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();

                            //dynamic ementaInserir = new ExpandoObject();
                            //ementaInserir.ListaArquivos = new List<ArquivoUpload>();
                            //itemListaVez.ListaEmenta = new List<dynamic>();

                            urlTratada = itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty);
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string dadosPdf = string.Empty;
                            string nomeArq = string.Empty;
                            //string htmlPage = string.Empty;
                            string corpoDados = string.Empty;

                            htmlPage = System.Net.WebUtility.HtmlDecode(getHtmlPaginaByGet(urlTratada, "default", string.Empty, "accept").Replace("\"", string.Empty).Replace("\t", " ").Replace("\r", " ").Replace("\n", " ").Replace("'", string.Empty));
                            htmlPage = Regex.Replace(htmlPage.Trim(), @"\s+", " ");

                            List<string> listMarcacao = new List<string>() { htmlPage.ToLower().IndexOf("<table border=0 cellspacing=0 cellpadding=0 width=100%").ToString() + "|</body>|<font face=arial size=3|</font>|<font face=times new roman size=2>|</font>|<font face=times new roman size=2 color=#008080|</font>" };

                            //,htmlPage.ToLower().IndexOf("").ToString() + "|||||||"

                            listMarcacao.RemoveAll(x => x.Contains("-1"));

                            listMarcacao = listMarcacao.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                            if (listMarcacao.Count != 1) listMarcacao = new List<string>();

                            if (htmlPage.Substring(int.Parse(listMarcacao[0].Split('|')[0])).ToLower().IndexOf(listMarcacao[0].Split('|')[1]) > 0)
                                corpoDados = htmlPage.Substring(int.Parse(listMarcacao[0].Split('|')[0]));
                            else
                                corpoDados = htmlPage.Substring(int.Parse(listMarcacao[0].Split('|')[0])).Substring(0, htmlPage.Substring(int.Parse(listMarcacao[0].Split('|')[0])).ToLower().IndexOf(listMarcacao[0].Split('|')[1]));

                            var listaDecretos = Regex.Split(corpoDados.ToLower(), "<font face=arial size=3").ToList();

                            listaDecretos.RemoveAt(0);

                            foreach (string itemDecretos in listaDecretos)
                            {
                                try
                                {
                                    dynamic ementaInserir = new ExpandoObject();
                                    ementaInserir.ListaArquivos = new List<ArquivoUpload>();
                                    itemListaVez.ListaEmenta = new List<dynamic>();

                                    numero = string.Empty;
                                    especie = string.Empty;
                                    titulo = string.Empty;
                                    publicacao = string.Empty;
                                    edicao = string.Empty;
                                    ementa = string.Empty;
                                    texto = string.Empty;
                                    dadosPdf = string.Empty;
                                    nomeArq = string.Empty;

                                    if (string.IsNullOrEmpty(titulo))
                                        titulo = itemDecretos.Substring(1).Substring(0, itemDecretos.Substring(1).ToLower().IndexOf(listMarcacao[0].Split('|')[3]));

                                    titulo = ObterStringLimpa(titulo);

                                    if (listMarcacao.Count > 0 && !listMarcacao[0].Split('|')[4].Equals("nd") && itemDecretos.ToLower().IndexOf(listMarcacao[0].Split('|')[4]) >= 0)
                                        ementa = itemDecretos.Substring(itemDecretos.ToLower().IndexOf(listMarcacao[0].Split('|')[4])).Substring(0, itemDecretos.Substring(itemDecretos.ToLower().IndexOf(listMarcacao[0].Split('|')[4])).ToLower().IndexOf(listMarcacao[0].Split('|')[5]));

                                    titulo = ObterStringLimpa(titulo);
                                    ementa = ObterStringLimpa(ementa);

                                    if (itemDecretos.ToLower().IndexOf(listMarcacao[0].Split('|')[6]) >= 0 && !listMarcacao[0].Split('|')[6].Equals("nd"))
                                        publicacao = itemDecretos.Substring(itemDecretos.ToLower().IndexOf(listMarcacao[0].Split('|')[6])).Substring(0, itemDecretos.Substring(itemDecretos.ToLower().IndexOf(listMarcacao[0].Split('|')[6])).ToLower().IndexOf(listMarcacao[0].Split('|')[7]));

                                    if (publicacao.Equals(string.Empty))
                                    {
                                        List<int> listPublicacao = new List<int>() { titulo.ToLower().IndexOf(","), titulo.ToLower().IndexOf(" de ") };

                                        listPublicacao.RemoveAll(x => x == -1);

                                        if (listPublicacao.Count > 0)
                                            publicacao = titulo.Substring(listPublicacao.Min(x => x) + 4);
                                    }

                                    publicacao = ObterStringLimpa(publicacao);

                                    titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                                    /*Captura CSS*/

                                    var listaCss = Regex.Split(htmlPage.ToLower(), "<link").ToList();

                                    listaCss.RemoveAt(0);

                                    listaCssTratada = new List<string>();

                                    string novaCss = string.Empty;

                                    listaCss.ForEach(delegate(string x)
                                    {
                                        novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                        if (!novaCss.ToLower().Contains("http"))
                                            novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                        listaCssTratada.Add(novaCss);
                                    });

                                    novaCss = string.Empty;

                                    listaCssTratada.ForEach(x => novaCss += x);

                                    string descStyle = string.Empty;

                                    if (htmlPage.ToLower().Contains("<style"))
                                        descStyle = htmlPage.Substring(htmlPage.ToLower().IndexOf("<style"))
                                                                   .Substring(0, htmlPage.Substring(htmlPage.ToLower().IndexOf("<style")).ToLower().IndexOf("</style>") + 8);

                                    /*Fim Captura CSS*/

                                    texto = novaCss + descStyle + itemDecretos;

                                    /** Outros **/
                                    ementaInserir.DescSigla = string.Empty;
                                    ementaInserir.HasContent = false;

                                    /** Arquivo **/
                                    ementaInserir.Tipo = 3;
                                    ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "CE") + "}";

                                    if (numero.Equals(string.Empty))
                                        numero = "0";

                                    especie = titulo.Substring(0, titulo.IndexOf(" "));

                                    /** Default **/
                                    ementaInserir.Publicacao = publicacao;
                                    ementaInserir.Sigla = "Sefaz CE";
                                    ementaInserir.Republicacao = string.Empty;
                                    ementaInserir.Ementa = ementa;
                                    ementaInserir.TituloAto = titulo;
                                    ementaInserir.Especie = especie;
                                    ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                                    ementaInserir.DataEdicao = edicao;
                                    ementaInserir.Texto = texto;
                                    ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(removeTagScript(texto))));
                                    ementaInserir.IdFila = itemLista_Nivel2.Id;

                                    ementaInserir.Escopo = "CE";
                                    ementaInserir.IdFila = itemLista_Nivel2.Id;

                                    itemListaVez.ListaEmenta.Add(ementaInserir);

                                    dynamic itemFonte = new ExpandoObject();

                                    itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                                }
                                catch (Exception ex)
                                {
                                    if (!htmlPage.Contains("Erro ao processar página, detalhes:"))
                                        MessageBox.Show("Erro - index:" + inferno.ToString() + "\n UrlInterna = " + itemDecretos);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!htmlPage.Contains("Erro ao processar página, detalhes:"))
                                MessageBox.Show("Erro - index:" + inferno.ToString());
                        }
                    }
                }
            }

            #endregion

            #endregion

            #region "Sefaz MA"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {
                try
                {
                    this.Text = "Fonte: Sefaz MA";

                    string htmlPage = string.Empty;

                    List<string> linksSefazMA = new List<string>() { "http://sistemas1.sefaz.ma.gov.br/portalsefaz/jsp/pagina/pagina.jsf?codigo=97"
                                                       ,"http://sistemas1.sefaz.ma.gov.br/portalsefaz/files?codigo=9590|f"
                                                       ,"http://sistemas1.sefaz.ma.gov.br/portalsefaz/jsp/pagina/pagina.jsf?codigo=99"
                                                       ,"http://sistemas1.sefaz.ma.gov.br/portalsefaz/jsp/pagina/pagina.jsf?codigo=101"
                                                       ,"http://sistemas1.sefaz.ma.gov.br/portalsefaz/jsp/pagina/pagina.jsf?codigo=104"
                                                       ,"http://sistemas1.sefaz.ma.gov.br/portalsefaz/jsp/pagina/pagina.jsf?codigo=105"
                                                       ,"http://sistemas1.sefaz.ma.gov.br/portalsefaz/jsp/pagina/pagina.jsf?codigo=106"
                                                       ,"http://sistemas1.sefaz.ma.gov.br/portalsefaz/jsp/pagina/pagina.jsf?codigo=3936"};

                    foreach (var itemMa in linksSefazMA)
                    {
                        try
                        {
                            objUrll = new ExpandoObject();

                            objUrll.Indexacao = "Secretaria do estado da Fazendo do Maranhão";
                            objUrll.Url = itemMa;

                            objUrll.Lista_Nivel2 = new List<dynamic>();

                            if (!itemMa.Contains("|f"))
                            {
                                request = (HttpWebRequest)WebRequest.Create(itemMa.Substring(0, itemMa.IndexOf("?")));

                                var postData = itemMa.Substring(itemMa.IndexOf("?") + 1);
                                var data = Encoding.ASCII.GetBytes(postData);

                                request.Method = "POST";
                                request.ContentType = "application/x-www-form-urlencoded";
                                request.ContentLength = data.Length;

                                using (var stream = request.GetRequestStream())
                                {
                                    stream.Write(data, 0, data.Length);
                                }

                                var response = (HttpWebResponse)request.GetResponse();
                                var responseString = new StreamReader(response.GetResponseStream(), Encoding.Default).ReadToEnd();

                                responseString = System.Net.WebUtility.HtmlDecode(responseString.Replace("\n", " ").Replace("\r", " ").Replace("\t", " ")).Replace("\"", string.Empty);
                                responseString = responseString.Substring(responseString.IndexOf("<div class=row-fluid>"));
                                responseString = responseString.Substring(0, responseString.IndexOf("<div class=footer"));

                                var listaUrlIndex = Regex.Split(responseString, "href=").ToList();
                                listaUrlIndex.RemoveAt(0);

                                foreach (var itemUrlInt in listaUrlIndex)
                                {
                                    var nextUrl = itemMa.Substring(0, itemMa.IndexOf(".br/") + 4) + itemUrlInt.Substring(0, itemUrlInt.IndexOf(">")).Replace("target=_blank", string.Empty).Replace("../../../", string.Empty);

                                    if (!nextUrl.Contains("pdf") && !nextUrl.Contains("files"))
                                    {
                                        try
                                        {
                                            string htmlUrls = getHtmlPaginaByGet(nextUrl, "default").Replace("\"", string.Empty).Replace("\t", " ").Replace("\r", " ").Replace("\n", " ").Replace("'", string.Empty);

                                            List<string> listaCorte = new List<string>() { htmlUrls.ToLower().IndexOf("<p><p><strong>anexos</strong>").ToString() + "|<input type=hidden" 
                                                                              ,htmlUrls.ToLower().IndexOf("<div class=titulo_noticia>").ToString() + "|<input type=hidden"};

                                            listaCorte.RemoveAll(x => x.Contains("-1"));

                                            htmlUrls = htmlUrls.Substring(int.Parse(listaCorte[0].Split('|')[0])).Substring(0, htmlUrls.Substring(int.Parse(listaCorte[0].Split('|')[0])).ToLower().IndexOf(listaCorte[0].Split('|')[1]));

                                            var listaIndexUrl = Regex.Split(htmlUrls, "href=").ToList();

                                            listaIndexUrl.RemoveAt(0);

                                            foreach (var itemIntUrl in listaIndexUrl)
                                            {
                                                if (itemIntUrl.Substring(0, 4).Contains("http"))
                                                    nextUrl = itemIntUrl.Substring(0, itemIntUrl.IndexOf(">")).Replace("target=_blank", string.Empty);
                                                else
                                                    nextUrl = itemMa.Substring(0, itemMa.IndexOf(".br/") + 4) + itemIntUrl.Substring(0, itemIntUrl.IndexOf(">")).Replace("target=_blank", string.Empty).Replace("../../../", string.Empty);

                                                objItenUrl = new ExpandoObject();

                                                objItenUrl.Url = nextUrl;
                                                objUrll.Lista_Nivel2.Add(objItenUrl);
                                            }
                                        }
                                        catch (Exception)
                                        {
                                        }
                                    }
                                    else
                                    {
                                        objItenUrl = new ExpandoObject();

                                        objItenUrl.Url = nextUrl;
                                        objUrll.Lista_Nivel2.Add(objItenUrl);
                                    }
                                }
                            }
                            else
                            {
                                objItenUrl = new ExpandoObject();

                                objItenUrl.Url = itemMa.Replace("|f", string.Empty);
                                objUrll.Lista_Nivel2.Add(objItenUrl);
                            }

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
                catch (Exception)
                {
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazMa"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazMa");

                string urlTratada = string.Empty;
                string htmlPage = string.Empty;

                int inferno = 0;

                var listaXxx = new List<string>();

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            if (inferno < 0)
                            {
                                inferno++;
                                continue;
                            }

                            inferno++;

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty);
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string corpoDados = string.Empty;

                            string nomeArq = "arqPdf" + inferno.ToString() + ".pdf";

                            nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                            using (WebClient webclient = new WebClient())
                            {
                                webclient.DownloadFile(urlTratada, nomeArq);
                            }

                            string dadosPdf = LeArquivo(nomeArq);

                            File.Delete(nomeArq);

                            List<string> listInformacoes = new List<string>() { (dadosPdf.IndexOf("DECRETO ") >= 0 && dadosPdf.Contains("DOE") ? dadosPdf.IndexOf("DECRETO ").ToString() : "-1") + "|\n|DOE|\n|nd"
                                                                               ,(dadosPdf.IndexOf("DECRETO ") >= 0 && dadosPdf.Contains("D.O.E") ? dadosPdf.IndexOf("DECRETO ").ToString() : "-1") + "|\n|D.O.E|\n|nd"

                                                                               ,dadosPdf.IndexOf("DECRETO N\no\n").ToString() + "|\n|nd|nd|nd"

                                                                               ,(dadosPdf.IndexOf("PORTARIA ") >= 0 && dadosPdf.Contains("DOE") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|DOE|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA ") >= 0 && !dadosPdf.Contains("DOE") && dadosPdf.Contains("São Luís (MA), ") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|São Luís (MA), |\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA ") >= 0 && !dadosPdf.Contains("DOE") && dadosPdf.Contains("São Luís,") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|São Luís,|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA ") >= 0 && !dadosPdf.Contains("DOE") && dadosPdf.Contains("SÃO LUÍS,") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|SÃO LUÍS,|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA ") >= 0 && !dadosPdf.Contains("DOE") && dadosPdf.Contains("Publicação: ") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|Publicação: |\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA ") >= 0 && !dadosPdf.Contains("DOE") && dadosPdf.Contains("Publicada em ") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|Publicada em |\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA ") >= 0 && !dadosPdf.Contains("DOE") && dadosPdf.Contains("Ed: ") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|Ed: |\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA ") >= 0 && !dadosPdf.Contains("DOE") && !dadosPdf.Contains("Ed: ") && !dadosPdf.Contains("Publicada em ") && !dadosPdf.Contains("Publicação: ") && !dadosPdf.Contains("SÃO LUÍS,") && !dadosPdf.Contains("São Luís,") && !dadosPdf.Contains("São Luís (MA), ") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|nd|nd|nd"

                                                                               ,dadosPdf.IndexOf("Convênio ICMS 126/98").ToString() + "|\n|nd|nd|nd"

                                                                               ,(dadosPdf.IndexOf("LEI N") >= 0 && dadosPdf.Contains("DOE") ? dadosPdf.IndexOf("LEI N").ToString() : "-1") + "|\n|DOE|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LEI N") >= 0 && dadosPdf.Contains("D.O.E") ? dadosPdf.IndexOf("LEI N").ToString() : "-1") + "|\n|D.O.E|\n|nd"

                                                                               ,dadosPdf.IndexOf("RESOLUÇÃO ADMINISTRATIVA").ToString() + "|\n|DOE|\n|nd"

                                                                               ,dadosPdf.IndexOf("ATO DECLARATÓRIO").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("ATO DECLARATÓRIO DE").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("Convênios ").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("DECRETO   N").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("DECRETO  N").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("DECRETO N").ToString() + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("ECRETO N") >= 0 && dadosPdf.Contains("(DOE ") ? dadosPdf.IndexOf("ECRETO N").ToString() : "-1") + "|\n|(DOE |\n|nd"
                                                                               ,dadosPdf.IndexOf("LEI N").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("INSTRUÇÃO    N").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("INSTRUÇÃO N").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("Instrução Normativa N").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("LEI   N").ToString() + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("LEI   N") >= 0 && dadosPdf.Contains("DOE ") ? dadosPdf.IndexOf("LEI   N").ToString() : "-1") + "|\n|DOE |\n|nd"
                                                                               ,(dadosPdf.IndexOf("LEI ") >= 0 && dadosPdf.Contains("Publicado no DOE ") ? dadosPdf.IndexOf("LEI ").ToString() : "-1") + "|\n|Publicado no DOE |\n|nd"
                                                                               ,dadosPdf.IndexOf("LEI nº").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("LEI.").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("MEDIDA PROVISÓRIA N").ToString() + "|\n|DOE |\n|nd"
                                                                               ,dadosPdf.IndexOf("PORT ARIA PGE N").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("Portaria  n").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("Portaria  N").ToString() + "|\n|DOE|\n|nd"
                                                                               ,dadosPdf.IndexOf("Portaria n").ToString() + "|\n|nd|nd|nd"
                                                                               ,dadosPdf.IndexOf("Portaria N").ToString() + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("Portaria n") >= 0 && dadosPdf.Contains("D.O.E.") ? dadosPdf.IndexOf("Portaria n").ToString() : "-1") + dadosPdf.IndexOf("Portaria n").ToString() + "|\n|D.O.E.|\n|nd"
                                                                               ,dadosPdf.IndexOf("RELATÓRIO DE  ATO DECLARATÓRIO ").ToString() + "|\n|nd|nd|nd"};

                            listInformacoes.RemoveAll(x => x.Contains("-1"));

                            listInformacoes = listInformacoes.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                            if (listInformacoes.Count == 0)
                            {
                                listaXxx.Add(dadosPdf.Substring(0, 400));
                                continue;
                            }

                            if (!listInformacoes[0].Split('|')[1].Equals("nd"))
                                titulo = dadosPdf.Substring(int.Parse(listInformacoes[0].Split('|')[0])).Replace("\no\n", string.Empty).Substring(0, dadosPdf.Substring(int.Parse(listInformacoes[0].Split('|')[0])).Replace("\no\n", string.Empty).IndexOf(listInformacoes[0].Split('|')[1]));

                            titulo = Regex.Replace(titulo.Trim(), @"\s+", " ");

                            if (!listInformacoes[0].Split('|')[2].Equals("nd"))
                                publicacao = dadosPdf.Substring(dadosPdf.IndexOf(listInformacoes[0].Split('|')[2])).Substring(0, dadosPdf.Substring(dadosPdf.IndexOf(listInformacoes[0].Split('|')[2])).IndexOf(listInformacoes[0].Split('|')[3]));
                            else if (titulo.Trim().ToLower().IndexOf(" de") >= 0)
                                publicacao = titulo.Trim().Substring(titulo.Trim().ToLower().IndexOf(" de") + 3).Trim();


                            if (!listInformacoes[0].Split('|')[4].Equals("nd"))
                                especie = listInformacoes[0].Split('|')[4];

                            else
                                especie = titulo.Substring(0, titulo.IndexOf(" "));

                            texto = dadosPdf;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "MA") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = "Sefaz MA";
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(removeTagScript(texto))));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "MA";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!htmlPage.Contains("Erro ao processar página, detalhes:"))
                                MessageBox.Show("Erro - index:" + inferno.ToString());
                        }
                    }

                    //string novoNcm = "TITULO\n";
                    //listaXxx.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
                    //File.WriteAllText(@"C:\Temp\NovasNCM.csv", novoNcm);
                }
            }

            #endregion

            #endregion

            #region "Sefaz MT"

            #region "Captura Url's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                this.Text = "Fonte: Sefaz MA";

                string htmlPage = string.Empty;

                List<string> linksSefazMt = new List<string>() { "http://www.sefaz.mt.gov.br/spl/portalpaginalegislacao?parametroCodigoDestaque=2#"
                                                   ,"http://www.sefaz.mt.gov.br/spl/portalpaginalegislacao?parametroCodigoDestaque=1"
                                                   ,"http://www.sefaz.mt.gov.br/spl/portalpaginalegislacao?parametroCodigoDestaque=3"
                                                   ,"http://www.sefaz.mt.gov.br/spl/portalpaginalegislacao?parametroCodigoDestaque=6"};

                foreach (var itemMt in linksSefazMt)
                {
                    try
                    {
                        objUrll = new ExpandoObject();

                        objUrll.Indexacao = "Secretaria do estado da Fazendo do Mato Grosso";
                        objUrll.Url = itemMt;

                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        string htmlListUrl = getHtmlPaginaByGet(itemMt, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                        htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<table cellpadding=0 cellspacing=0 border=0 align=center")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<table cellpadding=0 cellspacing=0 border=0 align=center")).IndexOf("</table>"));

                        var listaUrlIndex = Regex.Split(htmlListUrl, "href=").ToList();
                        listaUrlIndex.RemoveAt(0);

                        foreach (var itemUrlInt in listaUrlIndex)
                        {
                            objItenUrl = new ExpandoObject();

                            objItenUrl.Url = itemUrlInt.Substring(0, itemUrlInt.IndexOf(" "));
                            objUrll.Lista_Nivel2.Add(objItenUrl);
                        }

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            #endregion

            #region "Captura Doc's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazMt"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazMt");

                string urlTratada = string.Empty;
                string htmlPage = string.Empty;

                int inferno = 0;

                var listaXxx = new List<string>();

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty);
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string corpoDados = string.Empty;

                            string htmlText = getHtmlPaginaByGet(urlTratada, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                            List<string> listInformacoes = new List<string>() { htmlText.IndexOf("<FONT SIZE=2 COLOR=336699 FACE=Verdana").ToString() + "|<FONT SIZE=4 FACE=Arial>|</FONT>|<FONT SIZE=2 COLOR=336699 FACE=Verdana>Ementa:|</FONT></B>|<FONT SIZE=2 FACE=Verdana>#6|</FONT>"
                                                                               /*,htmlText.IndexOf("<FONT SIZE=2 COLOR=336699 FACE=Verdana").ToString() + "|<FONT SIZE=4 FACE=Arial>PORTARIA|</FONT>|<FONT SIZE=2 COLOR=336699 FACE=Verdana>Ementa:|</FONT></B>|<FONT SIZE=2 FACE=Verdana>#6|</FONT>"*/};

                            listInformacoes.RemoveAll(x => x.Contains("-1"));
                            listInformacoes = listInformacoes.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                            if (listInformacoes.Count == 0)
                            {
                                listaXxx.Add(urlTratada);
                                continue;
                            }

                            texto = htmlText.Substring(int.Parse(listInformacoes[0].Split('|')[0]));

                            if (!listInformacoes[0].Split('|')[1].Equals("nd"))
                                titulo = htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[1])).Substring(0, htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[1])).IndexOf(listInformacoes[0].Split('|')[2]));

                            titulo = Regex.Replace(titulo.Trim(), @"\s+", " ");
                            titulo = ObterStringLimpa(titulo);

                            if (!listInformacoes[0].Split('|')[5].Equals("nd"))
                            {
                                if (listInformacoes[0].Split('|')[5].Contains("#"))
                                {
                                    publicacao = Regex.Split(htmlText, listInformacoes[0].Split('|')[5].Split('#')[0])[int.Parse(listInformacoes[0].Split('|')[5].Split('#')[1])];
                                }
                                else
                                    publicacao = htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[5])).Substring(0, htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[5])).IndexOf(listInformacoes[0].Split('|')[6]));
                            }
                            else if (titulo.Trim().ToLower().IndexOf(" de") >= 0)
                                publicacao = titulo.Trim().Substring(titulo.Trim().ToLower().IndexOf(" de") + 3).Trim();

                            if (!listInformacoes[0].Split('|')[3].Equals("nd"))
                                ementa = htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[3])).Substring(0, htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[3])).IndexOf(listInformacoes[0].Split('|')[4]));

                            publicacao = ObterStringLimpa(publicacao);
                            especie = titulo.Substring(0, titulo.IndexOf(" "));
                            ementa = ObterStringLimpa(ementa);

                            titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            /*Captura CSS*/

                            var listaCss = Regex.Split(htmlText.ToLower(), "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (htmlText.ToLower().Contains("<style"))
                                descStyle = htmlText.Substring(htmlText.ToLower().IndexOf("<style"))
                                                           .Substring(0, htmlText.Substring(htmlText.ToLower().IndexOf("<style")).ToLower().IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "MT") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = "Sefaz MT";
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = descStyle + novaCss + texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(removeTagScript(texto))));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "MT";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!htmlPage.Contains("Erro ao processar página, detalhes:"))
                                MessageBox.Show("Erro - index:" + inferno.ToString());
                        }
                    }

                    //string novoNcm = "TITULO\n";
                    //listaXxx.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
                    //File.WriteAllText(@"C:\Temp\NovasNCM.csv", novoNcm);
                }
            }

            #endregion

            #endregion

            #region "Sefaz MS"

            #region "Captura Url's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                this.Text = "Fonte: Sefaz MS";

                string htmlPage = string.Empty;

                List<string> linksSefazMS = new List<string>() {"http://aacpdappls.net.ms.gov.br/appls/legislacao/serc/legato.nsf/Decreto?OpenView&Start=1&Count=30&Collapse=2#2"
                                                                ,"http://aacpdappls.net.ms.gov.br/appls/legislacao/serc/legato.nsf/Decretos%20Legislativos?OpenView"
                                                                ,"http://aacpdappls.net.ms.gov.br/appls/legislacao/serc/legato.nsf/Instru%C3%A7%C3%A3o%20Normativa?OpenView"
                                                                ,"http://aacpdappls.net.ms.gov.br/appls/legislacao/serc/legato.nsf/Portarias?OpenView"
                                                                ,"http://aacpdappls.net.ms.gov.br/appls/legislacao/serc/legato.nsf/Resolu%C3%A7%C3%B5es?OpenView&Start=1&Count=30&Collapse=5#5"
                                                                ,"http://aacpdappls.net.ms.gov.br/appls/legislacao/serc/legato.nsf/7a2675fdf26e910204256b1f005348a7/d3cc39d3a6aeeda803256cc20066f1fb?OpenDocument#CAP%C3%8DTULO%20I"};

                foreach (var itemMs in linksSefazMS)
                {
                    try
                    {
                        objUrll = new ExpandoObject();

                        objUrll.Indexacao = "Secretaria do estado da Fazendo do Mato Grosso do Sul";
                        objUrll.Url = itemMs;

                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        string htmlListUrl = getHtmlPaginaByGet(itemMs, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                        htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<table border=0 cellpadding=2 cellspacing=0>")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<table border=0 cellpadding=2 cellspacing=0>")).IndexOf("<img src=/appls/legislacao/serc/legato.nsf/vwicn103.gif?OpenImageResource"));

                        var listaUrlIndex = Regex.Split(htmlListUrl, "href=").ToList();
                        listaUrlIndex.RemoveAt(0);
                        listaUrlIndex.RemoveAll(x => x.Trim().Equals(">"));

                        foreach (var itemUrlInt in listaUrlIndex)
                        {
                            string novaUrlDocs = string.Empty;

                            if (itemMs.Contains("?OpenView&Start=1&Count=30&Collapse=5#5"))
                            {
                                novaUrlDocs = (itemMs.Substring(0, itemMs.IndexOf(".br/") + 3) + itemUrlInt.Substring(0, itemUrlInt.IndexOf(" "))).Replace("amp;", string.Empty);

                                htmlListUrl = getHtmlPaginaByGet(novaUrlDocs, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");
                                htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<img src=/icons/collapse.gif")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<img src=/icons/collapse.gif")).IndexOf("<img src=/appls/legislacao/serc/legato.nsf/vwicn103.gif"));

                                var listaUrlIndex1 = Regex.Split(htmlListUrl, "href=").ToList();

                                listaUrlIndex1.RemoveAt(0);
                                listaUrlIndex1.RemoveAt(listaUrlIndex1.Count - 1);
                                listaUrlIndex1.RemoveAll(x => x.Trim().Equals(">"));
                                //listaUrlIndex1.RemoveAll(x => x.Contains("?OpenView&"));

                                foreach (var itemUrlDoc in listaUrlIndex1)
                                {
                                    try
                                    {
                                        novaUrlDocs = (itemMs.Substring(0, itemMs.IndexOf(".br/") + 3) + itemUrlDoc.Substring(0, itemUrlDoc.IndexOf(" "))).Replace("amp;", string.Empty).Replace("Count=30", "Count=300");

                                        htmlListUrl = getHtmlPaginaByGet(novaUrlDocs, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");
                                        htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<img src=/icons/collapse.gif")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<img src=/icons/collapse.gif")).IndexOf("<img src=/appls/legislacao/serc/legato.nsf/vwicn103.gif"));

                                        var listaUrlIndex2 = Regex.Split(htmlListUrl, "href=").ToList();

                                        listaUrlIndex2.RemoveAt(0);
                                        listaUrlIndex2.RemoveAt(listaUrlIndex2.Count - 1);
                                        listaUrlIndex2.RemoveAll(x => x.Contains("?OpenView&"));
                                        listaUrlIndex2.RemoveAll(x => x.Trim().Equals(">"));

                                        foreach (var item in listaUrlIndex2)
                                        {
                                            try
                                            {
                                                objItenUrl = new ExpandoObject();

                                                objItenUrl.Url = itemMs.Substring(0, itemMs.IndexOf(".br/") + 3) + item.Substring(0, item.IndexOf(">"));
                                                objUrll.Lista_Nivel2.Add(objItenUrl);
                                            }
                                            catch (Exception)
                                            {
                                            }
                                        }
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }
                            }
                            else
                            {
                                novaUrlDocs = (itemMs.Substring(0, itemMs.IndexOf(".br/") + 3) + itemUrlInt.Substring(0, itemUrlInt.IndexOf(" "))).Replace("amp;", string.Empty).Replace("Count=30", "Count=300");

                                htmlListUrl = getHtmlPaginaByGet(novaUrlDocs, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");
                                htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<img src=/icons/collapse.gif")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<img src=/icons/collapse.gif")).IndexOf("<img src=/appls/legislacao/serc/legato.nsf/vwicn103.gif"));

                                var listaUrlIndex1 = Regex.Split(htmlListUrl, "href=").ToList();

                                listaUrlIndex1.RemoveAt(0);
                                listaUrlIndex1.RemoveAt(listaUrlIndex1.Count - 1);
                                listaUrlIndex1.RemoveAll(x => x.Contains("?OpenView&"));

                                foreach (var itemUrlDoc in listaUrlIndex1)
                                {
                                    objItenUrl = new ExpandoObject();

                                    objItenUrl.Url = itemMs.Substring(0, itemMs.IndexOf(".br/") + 3) + itemUrlDoc.Substring(0, itemUrlDoc.IndexOf(">"));
                                    objUrll.Lista_Nivel2.Add(objItenUrl);
                                }
                            }
                        }

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            #endregion

            #region "Captura Doc's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazMs"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazMs");

                string urlTratada = string.Empty;
                string htmlPage = string.Empty;

                int inferno = 0;

                var listaXxx = new List<string>();

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty);
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string corpoDados = string.Empty;

                            string htmlText = getHtmlPaginaByGet(urlTratada, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                            List<string> listInformacoes = new List<string>() { (htmlText.IndexOf("<font size=2 face=Tahoma>") >= 0 && htmlText.IndexOf("<input name=Ato_DspNumDO type=hidden value=") >= 0 ? htmlText.IndexOf("<font size=2 face=Tahoma>").ToString() : "-1") + "|<input name=Ato_DspNum type=hidden value=|>|<input name=Ato_DspEmenta type=hidden value=|>|<input name=Ato_DspNumDO type=hidden value=|>"
                                                                               ,(htmlText.IndexOf("<font size=2 face=Tahoma>") >= 0 && htmlText.IndexOf("<input name=Ato_DspNumDO type=hidden value=") < 0 && htmlText.IndexOf("<b><font face=Verdana>") >= 0 ? htmlText.IndexOf("<font size=2 face=Tahoma>").ToString() : "-1") + "|<font size=2 face=Tahoma>|</font>|<i><font size=2 face=Tahoma>|</font>|<font face=Verdana>|</font>"
                                                                               ,(htmlText.IndexOf("<font size=2 face=Tahoma>") >= 0 && htmlText.IndexOf("<input name=Ato_DspNumDO type=hidden value=") < 0 && htmlText.IndexOf("<font face=Tahoma>") >= 0 && htmlText.IndexOf("<font size=2 face=Tahoma>Publicad") < 0 ? htmlText.IndexOf("<font size=2 face=Tahoma>").ToString() : "-1") + "|<font size=2 face=Tahoma>|</font>|<i><font size=2 face=Tahoma>|</font>|<font face=Tahoma>|</font>"
                                                                               ,(htmlText.IndexOf("<font size=2 face=Tahoma>") >= 0 && htmlText.IndexOf("<input name=Ato_DspNumDO type=hidden value=") < 0 && htmlText.IndexOf("<b><font face=Times New Roman>") >= 0 ? htmlText.IndexOf("<font size=2 face=Tahoma>").ToString() : "-1") + "|<font size=2 face=Tahoma>|</font>|<i><font size=2 face=Tahoma>|</font>|<b><font face=Times New Roman>|</font>"
                                                                               ,(htmlText.IndexOf("<font size=2 face=Tahoma>") >= 0 && htmlText.IndexOf("<input name=Ato_DspNumDO type=hidden value=") < 0 && htmlText.IndexOf("<b><font size=2 face=Tahoma><font size=4>") >= 0 ? htmlText.IndexOf("<font size=2 face=Tahoma>").ToString() : "-1") + "|<font size=2 face=Tahoma>|</font>|<i><font size=2 face=Tahoma>|</font>|<b><font size=2 face=Tahoma><font size=4>|</font>"
                                                                               ,(htmlText.IndexOf("<font size=2 face=Tahoma>") >= 0 && htmlText.IndexOf("<input name=Ato_DspNumDO type=hidden value=") < 0 && htmlText.IndexOf("<font size=2 face=Tahoma>Publicad") >= 0 ? htmlText.IndexOf("<font size=2 face=Tahoma>").ToString() : "-1") + "|<font size=2 face=Tahoma>|</font>|<i><font size=2 face=Tahoma>|</font>|<font size=2 face=Tahoma>Publicad|</font>"
                                                                               ,(htmlText.IndexOf("<font size=2 face=Tahoma>") >= 0 && htmlText.IndexOf("<input name=Ato_DspNumDO type=hidden value=") < 0 && htmlText.IndexOf("<font size=4 face=Times New Roman>") >= 0 ? htmlText.IndexOf("<font size=2 face=Tahoma>").ToString() : "-1") + "|<font size=2 face=Tahoma>|</font>|<i><font size=2 face=Tahoma>|</font>|<font size=4 face=Times New Roman>|</font>"
                                                                               ,(htmlText.IndexOf("<font size=2 face=Tahoma>") >= 0 && htmlText.IndexOf("<input name=Ato_DspNumDO type=hidden value=") < 0 && htmlText.IndexOf("<font size=2 face=Tahoma><b>") >= 0 ? htmlText.IndexOf("<font size=2 face=Tahoma>").ToString() : "-1") + "|<font size=2 face=Tahoma>|</font>|<i><font size=2 face=Tahoma>|</font>|<font size=2 face=Tahoma><b>|</font>"
                                                                               ,(htmlText.IndexOf("<font size=2 face=Tahoma>") >= 0 && 
                                                                                    htmlText.IndexOf("<input name=Ato_DspNumDO type=hidden value=") < 0 &&
                                                                                        htmlText.IndexOf("<font size=2 face=Tahoma><b>") < 0 && 
                                                                                            htmlText.IndexOf("<font size=4 face=Times New Roman>") < 0 &&
                                                                                                htmlText.IndexOf("<font size=2 face=Tahoma>Publicada") < 0 && 
                                                                                                    htmlText.IndexOf("<b><font size=2 face=Tahoma><font size=4>") < 0 && 
                                                                                                        htmlText.IndexOf("<b><font face=Times New Roman>") < 0 &&
                                                                                                            htmlText.IndexOf("<font face=Tahoma>") < 0 && 
                                                                                                                htmlText.IndexOf("<b><font face=Verdana>") < 0 ? htmlText.IndexOf("<font size=2 face=Tahoma>").ToString() : "-1") + "|<font size=2 face=Tahoma>|</font>|<i><font size=2 face=Tahoma>|</font>|nd|nd"};

                            listInformacoes.RemoveAll(x => x.Contains("-1"));
                            listInformacoes = listInformacoes.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                            if (listInformacoes.Count == 0)
                            {
                                listaXxx.Add(urlTratada);
                                continue;
                            }

                            texto = htmlText.Substring(int.Parse(listInformacoes[0].Split('|')[0]) + 25);

                            if (!listInformacoes[0].Split('|')[1].Equals("nd"))
                                titulo = texto.Substring(texto.IndexOf(listInformacoes[0].Split('|')[1])).Substring(0, texto.Substring(texto.IndexOf(listInformacoes[0].Split('|')[1])).IndexOf(listInformacoes[0].Split('|')[2]));

                            titulo = Regex.Replace(titulo.Trim(), @"\s+", " ").Replace("<input name=Ato_DspNum type=hidden value=", "<input name=Ato_DspNum type=hidden value=>");
                            titulo = ObterStringLimpa(titulo);

                            if (!listInformacoes[0].Split('|')[5].Equals("nd"))
                                publicacao = texto.Substring(texto.IndexOf(listInformacoes[0].Split('|')[5])).Substring(0, texto.Substring(texto.IndexOf(listInformacoes[0].Split('|')[5])).IndexOf(listInformacoes[0].Split('|')[6]));

                            else if (titulo.Trim().ToLower().IndexOf(" de") >= 0)
                                publicacao = titulo.Trim().Substring(titulo.Trim().ToLower().IndexOf(" de") + 3).Trim();

                            if (!listInformacoes[0].Split('|')[3].Equals("nd") && texto.IndexOf(listInformacoes[0].Split('|')[3]) >= 0)
                                ementa = texto.Substring(texto.IndexOf(listInformacoes[0].Split('|')[3])).Substring(0, texto.Substring(texto.IndexOf(listInformacoes[0].Split('|')[3])).IndexOf(listInformacoes[0].Split('|')[4]));

                            publicacao = ObterStringLimpa(publicacao.Replace("<input name=Ato_DspNumDO type=hidden value=", "<input name=Ato_DspNumDO type=hidden value=>"));
                            ementa = ObterStringLimpa(ementa.Replace("<input name=Ato_DspEmenta type=hidden value=", "<input name=Ato_DspEmenta type=hidden value=>"));

                            especie = titulo.Substring(0, titulo.IndexOf(" "));

                            titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            /*Captura CSS*/

                            var listaCss = Regex.Split(htmlText.ToLower(), "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (htmlText.ToLower().Contains("<style"))
                                descStyle = htmlText.Substring(htmlText.ToLower().IndexOf("<style"))
                                                           .Substring(0, htmlText.Substring(htmlText.ToLower().IndexOf("<style")).ToLower().IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "MS") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = "Sefaz MS";
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = descStyle + novaCss + texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(removeTagScript(texto))));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "MS";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            if (!htmlPage.Contains("Erro ao processar página, detalhes:"))
                                MessageBox.Show("Erro - index:" + inferno.ToString());
                        }
                    }

                    string novoNcm = "TITULO\n";
                    listaXxx.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
                    File.WriteAllText(@"C:\Temp\NovasNCM.csv", novoNcm);
                }
            }

            #endregion

            #endregion

            #region "Sefaz PB"

            #region "Captura Url's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                this.Text = "Fonte: Sefaz PB";

                string htmlPage = string.Empty;

                List<string> linksSefazPb = new List<string>() { "https://www.receita.pb.gov.br/ser/legislacao/34-leis"
                                                                ,"https://www.receita.pb.gov.br/ser/legislacao/35-decretos-estaduais"
                                                                ,"https://www.receita.pb.gov.br/ser/legislacao/193-regulamentos/ipva"
                                                                ,"https://www.receita.pb.gov.br/ser/legislacao/98-regulamentos/icms"
                                                                ,"https://www.receita.pb.gov.br/ser/legislacao/37-medidas-provisorias"
                                                                ,"https://www.receita.pb.gov.br/ser/legislacao/43-portarias"
                                                                ,"https://www.receita.pb.gov.br/ser/legislacao/135-instrucao-normativa"
                                                                ,"https://www.receita.pb.gov.br/ser/legislacao/197-leis/leis-estaduais/67-leis-estaduais"};

                foreach (var itemPb in linksSefazPb)
                {
                    try
                    {
                        objUrll = new ExpandoObject();

                        objUrll.Indexacao = "Secretaria do estado da Fazendo da Paraíba";
                        objUrll.Url = itemPb;

                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        string htmlListUrl = getHtmlPaginaByGet(itemPb, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                        htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<div class=items-container>")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<div class=items-container>")).IndexOf("<div id=syndicate class=mob-noNav>"));

                        var listaUrlIndex = Regex.Split(htmlListUrl, "<h3 class=item-title style=font-size:1.1em;").ToList();
                        listaUrlIndex.RemoveAt(0);

                        foreach (var itemUrlInt in listaUrlIndex)
                        {
                            objItenUrl = new ExpandoObject();

                            objItenUrl.Url = itemPb.Substring(0, itemPb.IndexOf(".br/") + 3) + itemUrlInt.Substring(itemUrlInt.IndexOf("href=") + 5).Substring(0, itemUrlInt.Substring(itemUrlInt.IndexOf("href=") + 5).IndexOf(" "));

                            if (!itemPb.Equals(objItenUrl.Url))
                                objUrll.Lista_Nivel2.Add(objItenUrl);
                        }

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazPb"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazPb");

                string urlTratada = string.Empty;
                string htmlPage = string.Empty;

                var listaXxx = new List<string>();

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty);
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string corpoDados = string.Empty;

                            string htmlText = getHtmlPaginaByGet(urlTratada, string.Empty).Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ").Replace("'", string.Empty);

                            List<string> listInformacoes = new List<string>() { (htmlText.IndexOf("<div class=content-text") >= 0 && htmlText.IndexOf("<span style=color: #ff0000;") >= 0 ? htmlText.IndexOf("<div class=content-text").ToString() : "-1") + "|<div class=attachmentsContainer>|<h2 id=content-title>|</h2>|<blockquote class=blockquote-reverse top-space|</blockquote>|<span style=color: #ff0000;|</span>" 
                                                                               ,(htmlText.IndexOf("<div class=content-text") >= 0 && htmlText.IndexOf("<span style=font-family: Arial")  >= 0 ? htmlText.IndexOf("<div class=content-text").ToString() : "-1") + "|<div class=attachmentsContainer>|<h2 id=content-title>|</h2>|<blockquote class=blockquote-reverse top-space|</blockquote>|<span style=font-family: Arial|</span>"
                                                                               ,(htmlText.IndexOf("<div class=content-text") >= 0 && htmlText.IndexOf("<b>PUBLICAD")  >= 0 ? htmlText.IndexOf("<div class=content-text").ToString() : "-1") + "|<div class=attachmentsContainer>|<h2 id=content-title>|</h2>|<blockquote class=blockquote-reverse top-space|</blockquote>|<b>PUBLICAD|</b>"
                                                                               ,(htmlText.IndexOf("<div class=content-text") >= 0 && htmlText.IndexOf("<b>PUBLICAD")  < 0 && htmlText.IndexOf("<span style=font-family: Arial")  < 0 && htmlText.IndexOf("<span style=color: #ff0000;")  < 0 ? htmlText.IndexOf("<div class=content-text").ToString() : "-1") + "|<div class=attachmentsContainer>|<h2 id=content-title>|</h2>|nd|nd|nd|nd"};

                            listInformacoes.RemoveAll(x => x.Contains("-1"));
                            listInformacoes = listInformacoes.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                            if (listInformacoes.Count == 0)
                            {
                                listaXxx.Add(urlTratada);
                                continue;
                            }

                            texto = htmlText.Substring(int.Parse(listInformacoes[0].Split('|')[0])).Substring(0, htmlText.Substring(int.Parse(listInformacoes[0].Split('|')[0])).IndexOf(listInformacoes[0].Split('|')[1]));

                            if (!listInformacoes[0].Split('|')[2].Equals("nd"))
                                titulo = htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[2])).Substring(0, htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[2])).IndexOf(listInformacoes[0].Split('|')[3]));

                            titulo = Regex.Replace(titulo.Trim(), @"\s+", " ");
                            titulo = System.Net.WebUtility.HtmlDecode(ObterStringLimpa(titulo));

                            if (!listInformacoes[0].Split('|')[6].Equals("nd"))
                                publicacao = htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[6])).Substring(0, htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[6])).IndexOf(listInformacoes[0].Split('|')[7]));

                            else if (titulo.Trim().ToLower().IndexOf(" de") >= 0)
                                publicacao = titulo.Trim().Substring(titulo.Trim().ToLower().IndexOf(" de") + 3).Trim();

                            if (!listInformacoes[0].Split('|')[4].Equals("nd") && htmlText.IndexOf(listInformacoes[0].Split('|')[4]) >= 0)
                                ementa = htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[4])).Substring(0, htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[4])).IndexOf(listInformacoes[0].Split('|')[5]));

                            publicacao = System.Net.WebUtility.HtmlDecode(ObterStringLimpa(publicacao));
                            ementa = System.Net.WebUtility.HtmlDecode(ObterStringLimpa(ementa));

                            especie = titulo.Substring(0, titulo.IndexOf(" "));

                            titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            /*Captura CSS*/

                            var listaCss = Regex.Split(htmlText.ToLower(), "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (htmlText.ToLower().Contains("<style"))
                                descStyle = htmlText.Substring(htmlText.ToLower().IndexOf("<style"))
                                                           .Substring(0, htmlText.Substring(htmlText.ToLower().IndexOf("<style")).ToLower().IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "PB") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = "Sefaz PB";
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = descStyle + novaCss + texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(removeTagScript(texto))));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "PB";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            listaXxx.Add(urlTratada);

                            //if (!htmlPage.Contains("Erro ao processar página, detalhes:"))
                            //    MessageBox.Show("Erro - index:" + inferno.ToString());
                        }
                    }

                    string novoNcm = "TITULO\n";
                    listaXxx.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
                    File.WriteAllText(@"C:\Temp\NovasNCM.csv", novoNcm);
                }
            }

            #endregion

            #endregion

            #region "SefazPA"

            #region "Captura Url's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                this.Text = "Fonte: Sefaz PA";

                string htmlPage = string.Empty;

                List<string> linksSefazPa = new List<string>() { "http://www.sefa.pa.gov.br/index.php/legislacao/estadual/10-leis-estaduais"
                                                                ,"http://www.sefa.pa.gov.br/legislacao/interna/decreto/dc2001_04676.pdf"
                                                                ,"http://www.sefa.pa.gov.br/legislacao/interna/decreto/dc2001_c4676.pdf"
                                                                ,"http://www.sefa.pa.gov.br/legislacao/interna/decreto/dc2006_02703.pdf"
                                                                ,"http://www.sefa.pa.gov.br/index.php/legislacao/estadual/2268-decretos"
                                                                ,"http://www.sefa.pa.gov.br/index.php/legislacao/3140-instrucoes-normativas|Por Ano de publica"
                                                                ,"http://www.sefa.pa.gov.br/index.php/legislacao/3130-portarias"
                                                                ,"http://www.sefa.pa.gov.br/index.php/legislacao/3135-resolucoes﻿|por Ano de Publica"};

                foreach (var itemPa in linksSefazPa)
                {
                    try
                    {
                        objUrll = new ExpandoObject();

                        objUrll.Indexacao = "Secretaria do estado da Fazendo da Pará";
                        objUrll.Url = itemPa;

                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        List<string> listaUrlIndex;
                        List<string> listaUrlIndexLvl2;

                        if (!itemPa.Contains(".pdf"))
                        {
                            string htmlListUrl = getHtmlPaginaByGet(itemPa.Split('|')[0], "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                            if (itemPa.Split('|').Length <= 1)
                                htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<span class=sub-category")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<span class=sub-category")).IndexOf("<hr class=separador-interno>"));
                            else
                                htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf(itemPa.Split('|')[1])).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf(itemPa.Split('|')[1])).IndexOf("<hr class=separador-interno>"));

                            listaUrlIndex = Regex.Split(htmlListUrl, "href=").ToList();
                            listaUrlIndex.RemoveAt(0);

                            listaUrlIndex = listaUrlIndex.Select(x => itemPa.Substring(0, itemPa.IndexOf(".br/") + 3) + x.Substring(0, x.IndexOf(">"))).ToList();

                            var listCopia = listaUrlIndex.Select(x => x).ToList();

                            foreach (var itemCopia in listCopia)
                            {
                                if (!itemCopia.Contains(".pdf"))
                                {
                                    listaUrlIndex.Remove(itemCopia);

                                    htmlListUrl = getHtmlPaginaByGet(itemCopia, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                                    htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<span class=sub-category")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<span class=sub-category")).IndexOf("<hr class=separador-interno>"));

                                    listaUrlIndexLvl2 = Regex.Split(htmlListUrl, "href=").ToList();
                                    listaUrlIndexLvl2.RemoveAt(0);

                                    foreach (var itemLvl2 in listaUrlIndexLvl2)
                                    {
                                        listaUrlIndex.Add(itemPa.Substring(0, itemPa.IndexOf(".br/") + 3) + itemLvl2.Substring(0, itemLvl2.IndexOf(">")));
                                    }
                                }
                            }
                        }
                        else
                        {
                            listaUrlIndex = new List<string>() { itemPa };
                        }

                        foreach (var itemUrlInt in listaUrlIndex)
                        {
                            objItenUrl = new ExpandoObject();
                            objItenUrl.Url = itemUrlInt;
                            objUrll.Lista_Nivel2.Add(objItenUrl);
                        }

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            #endregion

            #region "Captura Doc's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazPa"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazPa");

                string urlTratada = string.Empty;
                string htmlPage = string.Empty;

                var listaXxx = new List<string>();

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string corpoDados = string.Empty;

                            string nomeArq = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1);

                            nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                            using (WebClient webclient = new WebClient())
                            {
                                webclient.DownloadFile(urlTratada, nomeArq);
                            }

                            string dadosPdf = LeArquivo(nomeArq);

                            File.Delete(nomeArq);

                            List<string> listInformacoes = new List<string>() { (dadosPdf.IndexOf("LEI N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("LEI N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LEI N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("LEI N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("INTRUÇÃO NORMATIVA N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("INTRUÇÃO NORMATIVA N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("INTRUÇÃO NORMATIVA N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("INTRUÇÃO NORMATIVA N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA CONJUNTA N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA CONJUNTA N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA CONJUNTA N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA CONJUNTA N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("PORTARIA N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("PORTARIA N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("RESOLUÇÃO/CONSAT N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("RESOLUÇÃO/CONSAT N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("RESOLUÇÃO/CONSAT N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("RESOLUÇÃO/CONSAT N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("RESOLUÇÃO N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("RESOLUÇÃO N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("RESOLUÇÃO N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("RESOLUÇÃO N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("RESOLUÇÃO CONSAT N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("RESOLUÇÃO CONSAT N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("RESOLUÇÃO CONSAT N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("RESOLUÇÃO CONSAT N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("L E I N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("L E I N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("L E I N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("L E I N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA CONJUNTA N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("PORTARIA CONJUNTA N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA CONJUNTA N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("PORTARIA CONJUNTA N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA DFI N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("PORTARIA DFI N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA DFI N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("PORTARIA DFI N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LEI COMPLEMENTAR N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("LEI COMPLEMENTAR N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LEI COMPLEMENTAR N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("LEI COMPLEMENTAR N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LE I N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("LE I N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LE I N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("LE I N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("DECRETO N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("DECRETO N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("DECRETO N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("DECRETO N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LEI COMPLEMENTAR N") >= 0 && dadosPdf.Contains("Publicada  no") ? dadosPdf.IndexOf("LEI COMPLEMENTAR N").ToString() : "-1") + "|\n|Publicada  no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LEI COMPLEMENTAR N") >= 0 && dadosPdf.Contains("Publicado  no") ? dadosPdf.IndexOf("LEI COMPLEMENTAR N").ToString() : "-1") + "|\n|Publicado  no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LEI COMPLEMENTAR N") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("LEI COMPLEMENTAR N").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("LEI COMPLEMENTAR N") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("LEI COMPLEMENTAR N").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA 0") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA 0") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|Publicado no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("REGIMENTO INTERNO  \n") >= 0 && dadosPdf.Contains("Publicada no") ? dadosPdf.IndexOf("REGIMENTO INTERNO  \n").ToString() : "-1") + "|\n|Publicada no|\n|nd"
                                                                               ,(dadosPdf.IndexOf("REGIMENTO INTERNO  \n") >= 0 && dadosPdf.Contains("Publicado no") ? dadosPdf.IndexOf("REGIMENTO INTERNO  \n").ToString() : "-1") + "|\n|Publicado no|\n|nd"};

                            listInformacoes.RemoveAll(x => x.Contains("-1"));

                            listInformacoes = listInformacoes.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                            if (listInformacoes.Count == 0)
                            {
                                listaXxx.Add(dadosPdf.Substring(0, 400));
                                continue;
                            }

                            if (!listInformacoes[0].Split('|')[1].Equals("nd"))
                                titulo = dadosPdf.Substring(int.Parse(listInformacoes[0].Split('|')[0])).Replace("\no\n", string.Empty).Substring(0, dadosPdf.Substring(int.Parse(listInformacoes[0].Split('|')[0])).Replace("\no\n", string.Empty).IndexOf(listInformacoes[0].Split('|')[1]));

                            titulo = Regex.Replace(titulo.Trim(), @"\s+", " ");

                            if (!listInformacoes[0].Split('|')[2].Equals("nd"))
                                publicacao = dadosPdf.Substring(dadosPdf.IndexOf(listInformacoes[0].Split('|')[2])).Substring(0, dadosPdf.Substring(dadosPdf.IndexOf(listInformacoes[0].Split('|')[2])).IndexOf(listInformacoes[0].Split('|')[3]));
                            else if (titulo.Trim().ToLower().IndexOf(" de") >= 0)
                                publicacao = titulo.Trim().Substring(titulo.Trim().ToLower().IndexOf(" de") + 3).Trim();


                            if (!listInformacoes[0].Split('|')[4].Equals("nd"))
                                especie = listInformacoes[0].Split('|')[4];

                            else
                                especie = titulo.Substring(0, titulo.IndexOf(" "));

                            titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            texto = dadosPdf;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "PA") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = "Sefaz PA";
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(removeTagScript(texto)));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "PA";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                        }
                    }

                    //string novoNcm = "TITULO\n";
                    //listaXxx.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
                    //File.WriteAllText(@"C:\Temp\NovasNCM.csv", novoNcm);
                }
            }

            #endregion

            #endregion

            #region "SefazPE"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                this.Text = "Fonte: Sefaz PE";

                string htmlPage = string.Empty;

                List<string> linksSefazPa = new List<string>() { "https://www.sefaz.pe.gov.br/Legislacao/Tributaria/Paginas/Legislacao-Tributaria-Estadual.aspx" };

                foreach (var itemPe in linksSefazPa)
                {
                    try
                    {
                        objUrll = new ExpandoObject();

                        objUrll.Indexacao = "Secretaria do estado da Fazendo da PE";
                        objUrll.Url = itemPe;

                        objUrll.Lista_Nivel2 = new List<dynamic>();

                        List<string> listaUrlIndex;
                        List<string> listaUrlIndexLvl2;

                        string htmlListUrl = getHtmlPaginaByGet(itemPe.Split('|')[0], "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                        htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("Normas publicadas pela SEFAZ/PE</div>")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("Normas publicadas pela SEFAZ/PE</div>")).IndexOf("<div id=footer class=ms-dialogHidden>"));

                        listaUrlIndex = Regex.Split(htmlListUrl, "href=").ToList();
                        listaUrlIndex.RemoveAt(0);

                        listaUrlIndex = listaUrlIndex.Select(x => itemPe.Substring(0, itemPe.IndexOf(".br/") + 3) + x.Substring(0, x.IndexOf(" "))).ToList();
                        listaUrlIndex = listaUrlIndex.Select(x => x.Substring(0, (x + ">").IndexOf(">"))).ToList();

                        foreach (var itemUrlInt in listaUrlIndex)
                        {
                            try
                            {
                                htmlListUrl = getHtmlPaginaByGet(itemUrlInt, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                                htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<body")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<body")).IndexOf("</body>"));

                                listaUrlIndexLvl2 = Regex.Split(htmlListUrl, "href=").ToList();
                                listaUrlIndexLvl2.RemoveAt(0);

                                foreach (var itemLvl2 in listaUrlIndexLvl2)
                                {
                                    objItenUrl = new ExpandoObject();
                                    objItenUrl.Url = itemLvl2.Substring(0, itemLvl2.IndexOf(" "));
                                    objUrll.Lista_Nivel2.Add(objItenUrl);
                                }
                            }
                            catch (Exception)
                            {
                            }
                        }

                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazPe"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazPe");

                string urlTratada = string.Empty;
                string htmlPage = string.Empty;

                bool validar = false;

                var listaXxx = new List<string>();

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {
                            Thread.Sleep(1000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim().Replace("|f", string.Empty);
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string corpoDados = string.Empty;

                            string htmlText = getHtmlPaginaByGet(urlTratada, "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ").Replace("'", string.Empty);

                            List<string> listInformacoes = new List<string>() { (htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<p class=A01TtuloNorma") >= 0 && htmlText.IndexOf("<p class=A02DataPublicacao") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=A01TtuloNorma|</p>|nd|nd|<p class=A02DataPublicacao|</p>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<p class=A03Ementa") >= 0 && htmlText.IndexOf("<p class=A02DataPublicacao") < 0 && htmlText.IndexOf("<p class=A01TtuloNorma") >= 0 && htmlText.IndexOf("DOE de ") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=A01TtuloNorma|</p>|<p class=A03Ementa|</p>|DOE de |</p>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("color:windowtext>Publicado") < 0 && htmlText.IndexOf("<p class=A02DataPublicacao") < 0 && htmlText.IndexOf("<p class=A01TtuloNorma") >= 0 && htmlText.IndexOf("color:windowtext>Publicado") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=A01TtuloNorma|</p>|<p class=A03Ementa|</p>|color:windowtext>Publicado|</p>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<h2>PORTARIA") >= 0 && htmlText.IndexOf("<p class=A02DataPublicacao") < 0 && htmlText.IndexOf("<p class=A01TtuloNorma") < 0 && htmlText.IndexOf("DOE.") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<h2>PORTARIA|</h2>|nd|nd|DOE.|</"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<h2>PORTARIA") >= 0 && htmlText.IndexOf("<span style=color:navy>") < 0 && htmlText.IndexOf("<p class=01TtuloNorma") < 0 && htmlText.IndexOf("<p class=A02DataPublicacao") < 0 && htmlText.IndexOf("<p class=A01TtuloNorma") < 0 && htmlText.IndexOf("DOE.") < 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<h2>PORTARIA|</h2>|nd|nd|nd|nd"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("color:windowtext>Publicado") < 0 && htmlText.IndexOf("<p class=01TtuloNorma") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=01TtuloNorma|</p>|nd|nd|nd|nd"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("color:windowtext>Publicado") < 0 && htmlText.IndexOf("<p class=A02DataPublicacao") < 0 && htmlText.IndexOf("<p class=A01TtuloNorma") >= 0 && htmlText.IndexOf("DOE. ") < 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=A01TtuloNorma|</p>|nd|nd|nd|nd"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<span style=color:navy>") >= 0 && htmlText.IndexOf("<p class=01TtuloNorma") < 0 && htmlText.IndexOf("<p class=A02DataPublicacao") < 0 && htmlText.IndexOf("<p class=A01TtuloNorma") < 0 && htmlText.IndexOf("DOE.") < 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<span style=color:navy>|</span>|<p class=MsoBodyTextIndent|</p>|color:windowtext>Publicado|</span>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<p class=A07DOE") >= 0 && htmlText.IndexOf("<p class=A01TtuloNorma") >= 0 && htmlText.IndexOf("<p class=A07DOE") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=A01TtuloNorma|</p>|nd|nd|<p class=A07DOE|</p>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<span style=color:#1F497D") >= 0 && htmlText.IndexOf("<p class=A07DOE") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<span style=color:#1F497D|</p>|nd|nd|<p class=A07DOE|</p>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<p class=A01TtuloNorma") >= 0 && htmlText.IndexOf("<p class=A07DOE") < 0 && htmlText.IndexOf("<p class=A03Ementa") < 0? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=A01TtuloNorma|</p>|nd|nd|DOE de|</"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<p class=DTtuloPrincipalAzul") >= 0 && htmlText.IndexOf("<p class=A07DOE") < 0 && htmlText.IndexOf("<p class=A03Ementa") < 0? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=DTtuloPrincipalAzul|</p>|nd|nd|nd|nd"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<p class=A04Artigo-Pargrafo") >= 0 && htmlText.IndexOf("<span style=font-size:8.0pt") >= 0 && htmlText.IndexOf("<b style=mso-bidi-font-weight:normal") >= 0 && htmlText.IndexOf("<p class=A02DataPublicacao") < 0 && htmlText.IndexOf("<p class=A03Ementa") < 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=A04Artigo-Pargrafo|</p>|<b style=mso-bidi-font-weight:normal|</b>|<span style=font-size:8.0pt|</span>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<span  style=font-size:12.0pt;mso-bidi-font-size:8.0pt;font-family:Verdana,sans-serif;") >= 0 && htmlText.IndexOf("<span style=font-size:9.0pt;mso-bidi-font-size:") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<span  style=font-size:12.0pt;mso-bidi-font-size:8.0pt;font-family:Verdana,sans-serif;|</span>|nd|nd|<span style=font-size:9.0pt;mso-bidi-font-size:|</span>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<span style=font-size:12.0pt;mso-bidi-font-size:11.0pt;") >= 0 && htmlText.IndexOf("<p class=MsoNormal align=center style=text-align:center") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<span style=font-size:12.0pt;mso-bidi-font-size:11.0pt;|</span>|nd|nd|<p class=MsoNormal align=center style=text-align:center|</p>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<p class=TtuloNormaAzul") >= 0 && htmlText.IndexOf("DOE de ") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<p class=TtuloNormaAzul|</p>|nd|nd|DOE de |</p>"
                                                                               ,(htmlText.IndexOf("<body ") >= 0 && htmlText.IndexOf("<span style=mso-bidi-font-size:8.0pt;") >= 0 && htmlText.IndexOf("<p class=A02DataPublicacao") >= 0 && htmlText.IndexOf("<p class=A03Ementa") >= 0 ? htmlText.IndexOf("<body ").ToString() : "-1") + "|</body>|<span style=mso-bidi-font-size:8.0pt;|</span>|<p class=A03Ementa|</p>|<p class=A02DataPublicacao|</p>"
                            };

                            listInformacoes.RemoveAll(x => x.Contains("-1"));
                            listInformacoes = listInformacoes.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                            if (urlTratada.Equals("https://www.sefaz.pe.gov.br/Legislacao/Tributaria/Documents/legislacao/Decretos/1999/DEC21644_99.htm"))
                                validar = true;

                            if (listInformacoes.Count == 0)
                            {
                                if (validar)
                                    urlTratada = urlTratada;

                                listaXxx.Add(urlTratada);
                                continue;
                            }

                            texto = htmlText.Substring(int.Parse(listInformacoes[0].Split('|')[0])).Substring(0, htmlText.Substring(int.Parse(listInformacoes[0].Split('|')[0])).IndexOf(listInformacoes[0].Split('|')[1]));

                            if (!listInformacoes[0].Split('|')[2].Equals("nd"))
                                titulo = htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[2])).Substring(0, htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[2])).IndexOf(listInformacoes[0].Split('|')[3]));

                            titulo = Regex.Replace(titulo.Trim(), @"\s+", " ");
                            titulo = System.Net.WebUtility.HtmlDecode(ObterStringLimpa(titulo));

                            if (!listInformacoes[0].Split('|')[6].Equals("nd"))
                                publicacao = htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[6])).Substring(0, htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[6])).IndexOf(listInformacoes[0].Split('|')[7]));

                            else if (titulo.Trim().ToLower().IndexOf(" de") >= 0)
                                publicacao = titulo.Trim().Substring(titulo.Trim().ToLower().IndexOf(" de") + 3).Trim();

                            if (!listInformacoes[0].Split('|')[4].Equals("nd") && htmlText.IndexOf(listInformacoes[0].Split('|')[4]) >= 0)
                                ementa = htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[4])).Substring(0, htmlText.Substring(htmlText.IndexOf(listInformacoes[0].Split('|')[4])).IndexOf(listInformacoes[0].Split('|')[5]));

                            publicacao = System.Net.WebUtility.HtmlDecode(ObterStringLimpa(publicacao));
                            ementa = System.Net.WebUtility.HtmlDecode(ObterStringLimpa(ementa));

                            especie = titulo.Substring(0, titulo.IndexOf(" "));

                            titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            /*Captura CSS*/

                            var listaCss = Regex.Split(htmlText.ToLower(), "<link").ToList();

                            listaCss.RemoveAt(0);

                            listaCssTratada = new List<string>();

                            string novaCss = string.Empty;

                            listaCss.ForEach(delegate(string x)
                            {
                                novaCss = string.Format("{0}{1}", "<link", x.Substring(0, x.IndexOf(">") + 1));

                                if (!novaCss.ToLower().Contains("http"))
                                    novaCss = novaCss.Insert(novaCss.IndexOf("href=") + 5, ("http://" + urlTratada.Substring(urlTratada.IndexOf("//") + 2).Substring(0, urlTratada.Substring(urlTratada.IndexOf("//") + 2).IndexOf("/"))));

                                listaCssTratada.Add(novaCss);
                            });

                            novaCss = string.Empty;

                            listaCssTratada.ForEach(x => novaCss += x);

                            string descStyle = string.Empty;

                            if (htmlText.ToLower().Contains("<style"))
                                descStyle = htmlText.Substring(htmlText.ToLower().IndexOf("<style"))
                                                           .Substring(0, htmlText.Substring(htmlText.ToLower().IndexOf("<style")).ToLower().IndexOf("</style>") + 8);

                            /*Fim Captura CSS*/

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "PE") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = "Sefaz PE";
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = descStyle + novaCss + texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(ObterStringLimpa(removeTagScript(texto))));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "PE";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                            listaXxx.Add(urlTratada);
                        }
                    }

                    string novoNcm = "TITULO\n";
                    listaXxx.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
                    File.WriteAllText(@"C:\Temp\NovasNCM.csv", novoNcm);
                }
            }

            #endregion

            #endregion

            #region "SefazPI"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u1"))
            {

                this.Text = "Fonte: Sefaz PE";

                string htmlPage = string.Empty;

                List<string> linksSefazPi = new List<string>() { "http://webas.sefaz.pi.gov.br/legislacao/atos-normativos/?resultsToSkip={0}"
                                                                ,"http://webas.sefaz.pi.gov.br/legislacao/comunicados/?resultsToSkip={0}"
                                                                ,"http://webas.sefaz.pi.gov.br/legislacao/decretos/?resultsToSkip={0}"
                                                                ,"http://webas.sefaz.pi.gov.br/legislacao/instrucao-normativa/?resultsToSkip={0}"
                                                                ,"http://webas.sefaz.pi.gov.br/legislacao/leis/?resultsToSkip={0}"
                                                                ,"http://webas.sefaz.pi.gov.br/legislacao/pareceres/?resultsToSkip={0}"
                                                                ,"http://webas.sefaz.pi.gov.br/legislacao/portarias/?resultsToSkip={0}"
                                                                ,"http://webas.sefaz.pi.gov.br/legislacao/ricms/?resultsToSkip={0}"};

                foreach (var itemPi in linksSefazPi)
                {
                    try
                    {
                        int paginacao = 0;

                        while (true)
                        {
                            objUrll = new ExpandoObject();

                            objUrll.Indexacao = "Secretaria do estado da Fazendo da Piauí";
                            objUrll.Url = itemPi;

                            objUrll.Lista_Nivel2 = new List<dynamic>();

                            List<string> listaUrlIndex;

                            string htmlListUrl = getHtmlPaginaByGet(string.Format(itemPi, paginacao.ToString()), "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                            if (htmlListUrl.Contains("Respondeu com status de 500 - Internal Error."))
                                break;
                            else
                                paginacao += 10;

                            htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("<section class=main-content")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("<section class=main-content")).IndexOf("<div class=pagination"));

                            listaUrlIndex = Regex.Split(htmlListUrl, "href=").ToList();
                            listaUrlIndex.RemoveAt(0);

                            listaUrlIndex = listaUrlIndex.Select(x => itemPi.Substring(0, itemPi.IndexOf(".br/") + 3) + x.Substring(0, x.IndexOf(" "))).ToList();
                            listaUrlIndex = listaUrlIndex.Select(x => x.Substring(0, (x + ">").IndexOf(">"))).ToList();
                            listaUrlIndex.RemoveAll(y => !y.Contains("?attach=true"));

                            foreach (var itemUrlInt in listaUrlIndex)
                            {
                                try
                                {
                                    objItenUrl = new ExpandoObject();
                                    objItenUrl.Url = itemUrlInt;
                                    objUrll.Lista_Nivel2.Add(objItenUrl);

                                }
                                catch (Exception)
                                {
                                }
                            }

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            #endregion

            #region "Captura DOC's"

            if ((modoProcessamento.Equals("f") || modoProcessamento.Equals("d")) && siglaFonteProcessamento.Equals("sefazPi"))
            {
                if (listaUrl.Count == 0)
                    listaUrl = new BuscaLegalDao().ObterUrlsParaProcessamento("sefazPi");

                string urlTratada = string.Empty;
                string htmlPage = string.Empty;

                var listaXxx = new List<string>();

                foreach (var nivel1_item in listaUrl)
                {
                    foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                    {
                        try
                        {

                            Thread.Sleep(2000);

                            dynamic itemListaVez = new ExpandoObject();
                            dynamic ementaInserir = new ExpandoObject();

                            ementaInserir.ListaArquivos = new List<ArquivoUpload>();

                            itemListaVez.ListaEmenta = new List<dynamic>();
                            urlTratada = itemLista_Nivel2.Url.Trim();
                            itemListaVez.Url = urlTratada;
                            itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                            string numero = string.Empty;
                            string especie = string.Empty;
                            string titulo = string.Empty;
                            string publicacao = string.Empty;
                            string edicao = string.Empty;
                            string ementa = string.Empty;
                            string texto = string.Empty;
                            string corpoDados = string.Empty;

                            urlTratada = urlTratada.Substring(0, urlTratada.IndexOf("?"));

                            string nomeArq = urlTratada.Substring(urlTratada.LastIndexOf("/") + 1);

                            nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), nomeArq);

                            using (WebClient webclient = new WebClient())
                            {
                                webclient.DownloadFile(urlTratada, nomeArq);
                            }

                            string dadosPdf = LeArquivo(nomeArq);

                            File.Delete(nomeArq);

                            List<string> listInformacoes = new List<string>() { (dadosPdf.IndexOf("ATO NORMATIVO") >= 0 && !dadosPdf.Contains("PUBLICADO NO DOE") ? dadosPdf.IndexOf("ATO NORMATIVO").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("ATO NORMATIVO") >= 0 && dadosPdf.Contains("PUBLICADO NO DOE") ? dadosPdf.IndexOf("ATO NORMATIVO").ToString() : "-1") + "|\n|PUBLICADO NO DOE|\n|nd"
                                                                               ,(dadosPdf.IndexOf("COMUNICADO ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("COMUNICADO ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("DECRETO N") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("DECRETO N").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("COMISSÃO TÉCNICA") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("COMISSÃO TÉCNICA").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("DECRETO  N") >= 0 && dadosPdf.Contains("Publicado no DOE") ? dadosPdf.IndexOf("DECRETO  N").ToString() : "-1") + "|\n|Publicado no DOE|\n|nd"
                                                                               ,(dadosPdf.IndexOf("DECRETO   N") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("DECRETO   N").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("LEI N") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("LEI N").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("PARECER ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("PARECER ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("PORTARIA ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("PORTARIA") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("PORTARIA").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("Portaria GSF") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("Portaria GSF").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("DECRETON") >= 0 && dadosPdf.Contains("Publicado no D.O.E") ? dadosPdf.IndexOf("DECRETON").ToString() : "-1") + "|\n|Publicado no D.O.E|\n|nd"
                                                                               ,(dadosPdf.IndexOf("DECRETO  N") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("DECRETO  N").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("INSTRUÇÃO NORMATIVA").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("LEI  N") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("LEI  N").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("PORTATRIA ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("PORTATRIA ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("ANISTIA ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("ANISTIA ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("Consulta sobre forma de  incidência de juros") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("Consulta sobre forma de  incidência de juros").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("NOTAS EXPLICATIVAS ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("NOTAS EXPLICATIVAS ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("LEI    ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("LEI    ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("INFORME SOBRE IPVA") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("INFORME SOBRE IPVA").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("MEMORANDO EXPORTAÇÃO N") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("MEMORANDO EXPORTAÇÃO N").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("LEI ORDINÁRIA N") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("LEI ORDINÁRIA N").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("DEPARTAMENTO DE ARRECADAÇÃO E TRIBUTAÇÃO") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("DEPARTAMENTO DE ARRECADAÇÃO E TRIBUTAÇÃO").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("CLASSIFICAÇÃO NACIONAL") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("CLASSIFICAÇÃO NACIONAL").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("Consulta sobre  procedimentos fiscais") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("Consulta sobre  procedimentos fiscais").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("LEI   N") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("LEI   N").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("DECRETO  ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("DECRETO  ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("Comunicado Sefaz ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("Comunicado Sefaz ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               ,(dadosPdf.IndexOf("ANEXO II – Art. 5º do Decreto ") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("ANEXO II – Art. 5º do Decreto ").ToString() : "-1") + "|\n|nd|nd|nd"
                                                                               /*,(dadosPdf.IndexOf("") >= 0 && dadosPdf.Contains("") ? dadosPdf.IndexOf("").ToString() : "-1") + "||||"*/
                                                                              };

                            listInformacoes.RemoveAll(x => x.Contains("-1"));
                            listInformacoes = listInformacoes.OrderBy(x => int.Parse(x.Split('|')[0])).ToList();

                            if (listInformacoes.Count == 0)
                            {
                                if (!string.Empty.Equals(dadosPdf.Trim()))
                                    listaXxx.Add(dadosPdf.Substring(0, 400));
                                continue;
                            }

                            if (!listInformacoes[0].Split('|')[1].Equals("nd"))
                                titulo = dadosPdf.Substring(int.Parse(listInformacoes[0].Split('|')[0])).Replace("\no\n", string.Empty).Substring(0, dadosPdf.Substring(int.Parse(listInformacoes[0].Split('|')[0])).Replace("\no\n", string.Empty).IndexOf(listInformacoes[0].Split('|')[1]));

                            titulo = Regex.Replace(titulo.Trim(), @"\s+", " ");

                            if (!listInformacoes[0].Split('|')[2].Equals("nd"))
                                publicacao = dadosPdf.Substring(dadosPdf.IndexOf(listInformacoes[0].Split('|')[2])).Substring(0, dadosPdf.Substring(dadosPdf.IndexOf(listInformacoes[0].Split('|')[2])).IndexOf(listInformacoes[0].Split('|')[3]));

                            else if (titulo.Trim().ToLower().IndexOf(" de") >= 0)
                                publicacao = titulo.Trim().Substring(titulo.Trim().ToLower().IndexOf(" de") + 3).Trim();


                            if (!listInformacoes[0].Split('|')[4].Equals("nd"))
                                especie = listInformacoes[0].Split('|')[4];

                            else
                                especie = (titulo + " ").Substring(0, (titulo + " ").IndexOf(" "));

                            titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                            texto = dadosPdf;

                            /** Outros **/
                            ementaInserir.DescSigla = string.Empty;
                            ementaInserir.HasContent = false;

                            /** Arquivo **/
                            ementaInserir.Tipo = 3;
                            ementaInserir.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "PI") + "}";

                            if (numero.Equals(string.Empty))
                                numero = "0";

                            /** Default **/
                            ementaInserir.Publicacao = publicacao;
                            ementaInserir.Sigla = "Sefaz PI";
                            ementaInserir.Republicacao = string.Empty;
                            ementaInserir.Ementa = ementa;
                            ementaInserir.TituloAto = titulo;
                            ementaInserir.Especie = especie;
                            ementaInserir.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                            ementaInserir.DataEdicao = edicao;
                            ementaInserir.Texto = texto;
                            ementaInserir.Hash = GerarHash(removerCaracterEspecial(removeTagScript(texto)));
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            ementaInserir.Escopo = "PI";
                            ementaInserir.IdFila = itemLista_Nivel2.Id;

                            itemListaVez.ListaEmenta.Add(ementaInserir);

                            dynamic itemFonte = new ExpandoObject();

                            itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });
                        }
                        catch (Exception ex)
                        {
                        }
                    }

                    //string novoNcm = "TITULO\n";
                    //listaXxx.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
                    //File.WriteAllText(@"C:\Temp\NovasNCM.csv", novoNcm);
                }
            }

            #endregion

            #endregion

            #region "Sefaz RO"

            #region "Captura URL's"

            if (modoProcessamento.Equals("f") || modoProcessamento.Equals("u"))
            {

                this.Text = "Fonte: Sefaz RO";

                webBrowser1 = new WebBrowser();

                webBrowser1.ScriptErrorsSuppressed = true;

                webBrowser1.Navigate(@"http://www.sefin.ro.gov.br/lista.jsp?tipo=lei&formato=108");

                webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompletedSefazRo);

/*              string htmlPage = string.Empty;

                List<string> linksSefazPi = new List<string>() { "http://www.sefin.ro.gov.br/lista.jsp?tipo=lei&formato=108&page={0}" };

                foreach (var itemPi in linksSefazPi)
                {
                    try
                    {
                        int paginacao = 1;

                        while (true)
                        {
                            objUrll = new ExpandoObject();

                            objUrll.Indexacao = "Secretaria do estado da Fazendo da Rondônia";
                            objUrll.Url = itemPi;

                            objUrll.Lista_Nivel2 = new List<dynamic>();

                            List<string> listaUrlIndex;

                            string htmlListUrl = getHtmlPaginaByGet(string.Format(itemPi, paginacao.ToString()), "default").Replace("\"", string.Empty).Replace("\n", " ").Replace("\t", " ").Replace("\r", " ");

                            if (!htmlListUrl.Contains("<span class=s13 gray"))
                                break;
                            else
                                paginacao++;

                            htmlListUrl = htmlListUrl.Substring(htmlListUrl.IndexOf("")).Substring(0, htmlListUrl.Substring(htmlListUrl.IndexOf("")).IndexOf(""));

                            listaUrlIndex = Regex.Split(htmlListUrl, "href=").ToList();
                            listaUrlIndex.RemoveAt(0);

                            listaUrlIndex = listaUrlIndex.Select(x => itemPi.Substring(0, itemPi.IndexOf(".br/") + 3) + x.Substring(0, x.IndexOf(" "))).ToList();
                            listaUrlIndex = listaUrlIndex.Select(x => x.Substring(0, (x + ">").IndexOf(">"))).ToList();

                            foreach (var itemUrlInt in listaUrlIndex)
                            {
                                try
                                {
                                    objItenUrl = new ExpandoObject();
                                    objItenUrl.Url = itemUrlInt;
                                    objUrll.Lista_Nivel2.Add(objItenUrl);
                                }
                                catch (Exception)
                                {
                                }
                            }

                            new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
 */
            }

            #endregion

            #endregion

        }

        private List<string> obterUrlNivelDocumento(string url)
        {
            Thread.Sleep(2000);

            WebRequest req = WebRequest.Create(url);
            WebResponse res = req.GetResponse();
            Stream dataStream = res.GetResponseStream();
            StreamReader reader;

            reader = new StreamReader(dataStream, Encoding.Default);

            string resposta = reader.ReadToEnd().Replace("\"", string.Empty).ToLower();

            var listRetorn = new List<string>();

            string contexto = resposta.Substring(resposta.LastIndexOf("<table width=100% cols=1 border=0")).Substring(0, resposta.Substring(resposta.LastIndexOf("<table width=100% cols=1 border=0")).IndexOf("</table>"));

            var listFolder = Regex.Split(contexto, "<tr nowrap").ToList();

            listFolder.RemoveAt(0);

            foreach (var item in listFolder)
            {
                if (item.Contains("href="))
                {
                    string novaUrl = item.Substring(item.IndexOf("href=") + 5).Substring(0, item.Substring(item.IndexOf("href=") + 5).IndexOf(" "));

                    if (item.Contains("/legislacaoonline/lpext.dll?f=images&fn=toc-collapsed.gif"))
                        listRetorn.AddRange(obterUrlNivelDocumento("http://www.sefaz.es.gov.br" + novaUrl));

                    if (item.Contains("/legislacaoonline/lpext.dll?f=images&fn=toc-leaf.gif"))
                        listRetorn.Add("http://www.sefaz.es.gov.br" + novaUrl);
                }

                if (url.Equals("http://www.sefaz.es.gov.br/LegislacaoOnline/lpext.dll?f=templates&fn=tools-contents.htm&cp=InfobaseLegislacaoOnline&2.0"))
                {
                    dynamic objUrll = new ExpandoObject();

                    objUrll.Indexacao = "Secretaria do estado da fazenda do ES";
                    objUrll.Url = item;
                    objUrll.Lista_Nivel2 = new List<dynamic>();

                    listRetorn.ForEach(delegate(string itemDado)
                    {
                        dynamic objItenUrl = new ExpandoObject();

                        objItenUrl.Url = itemDado;
                        objUrll.Lista_Nivel2.Add(objItenUrl);
                    });

                    new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { objUrll });

                    listRetorn = new List<string>();
                }
            }

            return listRetorn;
        }

        private List<string> obterUrlNivelDocumentoSefazMG(string url, string raiz)
        {
            Thread.Sleep(2000);

            string resposta = getHtmlPaginaByGet(url, "default").Replace("\"", string.Empty).ToLower();

            var listRetorn = new List<string>();

            string contexto = string.Empty;
            List<string> listFolder = new List<string>();

            if (resposta.IndexOf("<table class=msonormaltable") > 0)
            {
                while (resposta.IndexOf("<table class=msonormaltable") > 0)
                {
                    contexto = resposta.Substring(resposta.IndexOf("<table class=msonormaltable")).Substring(0, resposta.Substring(resposta.IndexOf("<table class=msonormaltable")).IndexOf("</table>"));
                    resposta = resposta.Remove(resposta.IndexOf("<table class=msonormaltable"), resposta.Substring(resposta.IndexOf("<table class=msonormaltable")).IndexOf("</table>"));

                    var listaAtual = Regex.Split(contexto, "href=").ToList();
                    listaAtual.RemoveAt(0);

                    listFolder.AddRange(listaAtual);
                }

                if (listFolder.Count > 0)
                    foreach (var item in listFolder)
                    {
                        string novaUrl = item.Substring(0, item.IndexOf(">"));

                        if (novaUrl.Contains("#"))
                            continue;

                        listRetorn.AddRange(obterUrlNivelDocumentoSefazMG(raiz + novaUrl, raiz));
                    }
                else
                    listRetorn.Add(url);

            }
            else
                listRetorn.Add(url);

            return listRetorn;
        }

        private List<string> obterUrlNivelDocumentoSefazRJ(string url, string raiz)
        {
            Thread.Sleep(2000);

            List<string> listKeys = new List<string>();

            string resposta = string.Empty;

            resposta = getHtmlPaginaByGet(url.Replace("|N", string.Empty).Replace("|s", string.Empty), "default").Replace("\"", string.Empty);

            listKeys.Add(resposta.ToLower().IndexOf("<table border=0 cellpadding=2 cellspacing=0").ToString() + "|</table>");
            listKeys.Add(resposta.ToLower().IndexOf("<div id=conteudosefaz").ToString() + "|</div>");

            listKeys.RemoveAll(x => x.Contains("-1"));
            var listRetorn = new List<string>();

            string contexto = string.Empty;
            List<string> listFolder = new List<string>();

            if (raiz.Contains("|N"))
                contexto = resposta.Substring(int.Parse(listKeys[0].Split('|')[0])).Substring(0, resposta.Substring(int.Parse(listKeys[0].Split('|')[0])).LastIndexOf(listKeys[0].Split('|')[1])).ToLower();
            else
                contexto = resposta.Substring(int.Parse(listKeys[0].Split('|')[0])).Substring(0, resposta.Substring(int.Parse(listKeys[0].Split('|')[0])).LastIndexOf(listKeys[0].Split('|')[1]));

            listFolder = Regex.Split(contexto, "href=").ToList();

            listFolder.RemoveAt(0);
            listFolder.RemoveAll(x => x.Contains("/constest.nsf/indiceint?openform"));
            listFolder.RemoveAll(x => x.Contains("/contlei.nsf/leicompint?openform"));
            listFolder.RemoveAll(x => x.Contains("/contlei.nsf/leiordint?openform"));

            string novaRaiz = string.Empty;

            if (raiz.Contains("|N"))
            {
                novaRaiz = string.Empty;

                var listNext = Regex.Split(resposta, "<area href=/").ToList();

                novaRaiz = raiz.Contains("/constest.nsf") ? raiz.Substring(0, raiz.IndexOf("/constest.nsf")) : novaRaiz;
                novaRaiz = raiz.Contains("/contlei.nsf") ? raiz.Substring(0, raiz.IndexOf("/contlei.nsf")) : novaRaiz;

                if (listFolder.Count > 0)
                    foreach (var item in listFolder)
                    {
                        string novaUrl = item.Substring(0, item.IndexOf(">"));

                        if (novaUrl.Contains("#"))
                            continue;

                        listRetorn.Add(novaRaiz + novaUrl);
                    }

                string urlPartial = listNext[listNext.Count - 2].Substring(0, listNext[listNext.Count - 2].IndexOf(" "));

                novaRaiz = novaRaiz + "/" + urlPartial.Replace("amp;", string.Empty);

                if (!url.Equals(novaRaiz))
                    listRetorn.AddRange(obterUrlNivelDocumentoSefazRJ(novaRaiz, raiz));
            }
            else if (listKeys.Count > 0 && raiz.Contains("|s"))
            {
                novaRaiz = raiz.Substring(0, raiz.IndexOf(".br/") + 3);

                if (listFolder.Count > 0)
                    foreach (var item in listFolder)
                    {
                        string novaUrl = item.Substring(0, item.IndexOf(" "));

                        if (novaUrl.Contains("#"))
                            continue;

                        if (!novaUrl.Contains(".br/"))
                            listRetorn.AddRange(obterUrlNivelDocumentoSefazRJ(novaUrl, raiz));
                        else
                            listRetorn.AddRange(obterUrlNivelDocumentoSefazRJ(novaRaiz + novaUrl, raiz));
                    }
                else
                    listRetorn.Add(url);
            }
            else
                listRetorn.Add(url);

            return listRetorn;
        }

        private string getHtmlPaginaByGet(string url, string encode, string cookie = "", string language = "")
        {
            string resposta = string.Empty;

            try
            {
                WebRequest reqNivel2_ = WebRequest.Create(url);

                if (!string.IsNullOrEmpty(cookie))
                    reqNivel2_.Headers.Add("Cookie", cookie);

                if (!string.IsNullOrEmpty(language))
                    reqNivel2_.Headers.Add("Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.6,en;q=0.4");

                WebResponse resNivel2_ = reqNivel2_.GetResponse();
                Stream dataStreamNivel2_ = resNivel2_.GetResponseStream();

                StreamReader readerNivel2_;

                if (encode.Equals("default"))
                    readerNivel2_ = new StreamReader(dataStreamNivel2_, Encoding.Default);
                else
                    readerNivel2_ = new StreamReader(dataStreamNivel2_);

                resposta = readerNivel2_.ReadToEnd();
            }
            catch (Exception ex)
            {
                resposta = "Erro ao processar página, detalhes: " + ex.Message;
            }

            return resposta;
        }

        private string obterDadosColection(string texto, string dado)
        {
            foreach (string item in Regex.Split(texto.ToLower(), "class=msonormal").ToList())
            {
                if (item.Contains(dado))
                {
                    return "<" + item;
                }
            }

            return string.Empty;
        }

        private string obterEmentaTexto(string texto)
        {
            string textoRetorno = string.Empty;

            var listaEmenta = Regex.Split(texto.ToLower(), "class=resoluoassunto").ToList();

            listaEmenta.RemoveAt(0);

            listaEmenta.ForEach(x => textoRetorno += "<" + x.Substring(0, x.IndexOf("</p>")));

            return textoRetorno;
        }

        private int obterPontoCorte(string titulo)
        {
            int ponto = 0;

            foreach (char x in titulo.ToList())
            {
                if (char.IsNumber(x))
                    return ponto;
                ponto++;
            }

            return ponto;
        }

        private dynamic obterEstruturaProcessarUrl_RfbLeg(string url)
        {

            dynamic objUrll = new ExpandoObject();

            try
            {
                WebRequest req = WebRequest.Create(url);
                WebResponse res = req.GetResponse();
                Stream dataStream = res.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);

                dynamic objItenUrl;

                objUrll.Indexacao = "Receita Federal";
                objUrll.Url = url;
                objUrll.Lista_Nivel2 = new List<dynamic>();

                string resposta = reader.ReadToEnd();

                resposta = System.Net.WebUtility.HtmlDecode(resposta.Replace("\n", string.Empty).Replace("\r", string.Empty).Replace("\t", string.Empty));

                var listaMPs = Regex.Split(resposta.Substring(resposta.IndexOf("invisible")).Replace("\"", string.Empty), "<tr>").ToList().FindAll(x => x.Contains("external-link"));

                listaMPs.RemoveAt(listaMPs.Count - 1);

                listaMPs.ForEach(delegate(string x)
                {
                    objItenUrl = new ExpandoObject();

                    objItenUrl.Url = x.Substring(x.IndexOf("href=") + 5, x.IndexOf("target") - (x.IndexOf("href=") + 5));

                    if (!objItenUrl.Url.Contains("www4"))
                        objUrll.Lista_Nivel2.Add(objItenUrl);
                });
            }
            catch (Exception ex)
            {
            }

            return objUrll;
        }

        public string obterTextoConteudo(string respostaNivel2_, int tipo)
        {
            int posicao = 0;

            int posicao_a = 0;
            int posicao_b = 0;
            int posicao_c = 0;

            string aux;
            string texto;

            if (tipo == 3 || tipo == 2)
            {
                posicao_a = respostaNivel2_.IndexOf("<div class=divSegmentos item>");
                posicao_b = respostaNivel2_.IndexOf("<div class=divSegmentos nao identificad>");
                posicao_c = respostaNivel2_.IndexOf("<div class=divSegmentos autor>");

                if (posicao_a >= 0 && (posicao_a < posicao_b || posicao_b < 0) && (posicao_a < posicao_c || posicao_c < 0))
                    posicao = posicao_a + "<div class=divSegmentos item>".Length;

                if (posicao_b >= 0 && (posicao_b < posicao_a || posicao_a < 0) && (posicao_b < posicao_c || posicao_c < 0))
                    posicao = posicao_b + "<div class=divSegmentos nao identificad>".Length;

                if (posicao_c >= 0 && (posicao_c < posicao_a || posicao_a < 0) && (posicao_c < posicao_b || posicao_b < 0))
                    posicao = posicao_c + "<div class=divSegmentos autor>".Length;

                if (respostaNivel2_.Contains("class=tright block"))
                {
                    texto = "<" + respostaNivel2_.Substring(respostaNivel2_.IndexOf("class=tright block")).Substring(0, respostaNivel2_.Substring(respostaNivel2_.IndexOf("class=tright block")).IndexOf("</a>"));
                }
                else
                {
                    aux = respostaNivel2_.Substring(posicao);

                    posicao = (aux.IndexOf("<div class=divSegmentos fecho>") >= 0 ? aux.IndexOf("<div class=divSegmentos fecho>") : aux.IndexOf("class=divSegmentos assinatura>"));

                    if (posicao >= 0)
                        aux = aux.Substring(0, posicao);

                    texto = string.Empty;

                    Regex.Split(aux, "</p>").ToList().ForEach(delegate(string x) { if (x.Contains("<p>")) { texto += x.Substring(x.IndexOf("<p>") + "<p>".Length); } });
                }

                bool key = false;
                aux = string.Empty;

                /*Isolar essa parte em uma nova função*/
                texto.ToList().ForEach(delegate(char x) { key = x.ToString().Equals("<") || (key && !x.ToString().Equals(">")); aux += !key && !x.ToString().Equals(">") ? x.ToString() : string.Empty; });

                texto = aux;
            }
            else
            {
                Boolean isFull = false;

                int posicao_corte = (respostaNivel2_.IndexOf("span class=ementa") < 0 ? 0 : respostaNivel2_.IndexOf("span class=ementa") + "span class=ementa".Length);

                posicao_a = respostaNivel2_.Substring(posicao_corte).IndexOf("div class=item>");

                posicao_b = respostaNivel2_.Substring(posicao_corte).IndexOf("<div class=nao identificad>");

                posicao_c = respostaNivel2_.Substring(posicao_corte).IndexOf("<div class=autor>");

                if (posicao_a >= 0 && (posicao_a < posicao_b || posicao_b < 0) && (posicao_a < posicao_c || posicao_c < 0))
                    posicao = posicao_a + "div class=item>".Length;

                if (posicao_b >= 0 && (posicao_b < posicao_a || posicao_a < 0) && (posicao_b < posicao_c || posicao_c < 0))
                    posicao = posicao_b + "<div class=nao identificad>".Length;

                if (posicao_c >= 0 && (posicao_c < posicao_a || posicao_a < 0) && (posicao_c < posicao_b || posicao_b < 0))
                {
                    posicao = posicao_c + "<div class=autor>".Length;

                    isFull = true;
                }

                aux = respostaNivel2_.Substring(posicao_corte).Substring(posicao);

                if (!isFull || aux.IndexOf("<div class=assinatura>") < 0)
                {
                    posicao = aux.IndexOf("<div class=fecho>");

                    if (posicao >= 0)
                        aux = aux.Substring(0, posicao);
                }
                else
                {
                    posicao = aux.IndexOf("<div class=assinatura>");

                    if (posicao >= 0)
                        aux = aux.Substring(0, posicao);
                }

                texto = string.Empty;

                bool key = false;

                /*Isolar essa parte em uma nova função*/
                aux.ToList().ForEach(delegate(char x) { key = x.ToString().Equals("<") || (key && !x.ToString().Equals(">")); texto += !key && !x.ToString().Equals(">") ? x.ToString() : string.Empty; });

                //texto = texto.Replace(">", string.Empty);
                //Regex.Split(aux, "</p>").ToList().ForEach(delegate(string x) { if (x.Contains("<p>")) { texto += x.Substring(x.IndexOf("<p>") + "<p>".Length); } });
            }

            return texto;
        }

        public string removerCaracterEspecial(string texto)
        {
            string antigoTexto = "ÄÅÁÂÀÃäáâàãÉÊËÈéêëèÍÎÏÌíîïìÖÓÔÒÕöóôòõÜÚÛüúûùÇç";
            string novoTexto = "AAAAAAaaaaaEEEEeeeeIIIIiiiiOOOOOoooooUUUuuuuCc";

            for (int i = 0; i < antigoTexto.Length; i++)
            {
                texto = texto.Replace(antigoTexto[i].ToString(), novoTexto[i].ToString());
            }
            return texto;
        }

        private string GerarHash(string valor)
        {
            using (MD5 hash = MD5.Create())
            {
                // Convert the input string to a byte array and compute the hash.
                byte[] data = hash.ComputeHash(Encoding.UTF8.GetBytes(valor));

                // Create a new Stringbuilder to collect the bytes
                // and create a string.
                StringBuilder sBuilder = new StringBuilder();

                // Loop through each byte of the hashed data
                // and format each one as a hexadecimal string.
                for (int i = 0; i < data.Length; i++)
                {
                    sBuilder.Append(data[i].ToString("x2"));
                }

                // Return the hexadecimal string.
                return sBuilder.ToString();
            }
        }

        public string LeArquivo(string fileName)
        {
            var text = new StringBuilder();

            // The PdfReader object implements IDisposable.Dispose, so you can
            // wrap it in the using keyword to automatically dispose of it
            using (var pdfReader = new PdfReader(fileName))
            {
                // Loop through each page of the document
                for (var page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                    var currentText = PdfTextExtractor.GetTextFromPage(
                        pdfReader,
                        page,
                        strategy);

                    currentText =
                        Encoding.UTF8.GetString(Encoding.Convert(
                            Encoding.Default,
                            Encoding.UTF8,
                            Encoding.Default.GetBytes(currentText)));

                    text.Append(currentText);
                }
            }

            return text.ToString();
        }

        private string ObterStringLimpa(string y)
        {
            bool key = false;
            string txtNovo = " ";

            y.ToList().ForEach(delegate(char x) { key = x.ToString().Equals("<") || (key && !x.ToString().Equals(">")); txtNovo += !key && !x.ToString().Equals(">") ? x.ToString() : string.Empty; });

            return Regex.Replace(txtNovo.Trim(), @"\s+", " ");
        }

        private string ObterStringLimpa(string y, string corte)
        {
            var listItens = Regex.Split(y, corte).ToList();

            bool key;
            string txtNovo = string.Empty;

            foreach (var item in listItens)
            {
                key = false;
                txtNovo = " ";

                item.ToList().ForEach(delegate(char x) { key = x.ToString().Equals("<") || (key && !x.ToString().Equals(">")); txtNovo += !key && !x.ToString().Equals(">") ? x.ToString() : string.Empty; });

                if (!txtNovo.Trim().Equals(string.Empty)) break;
            }

            return Regex.Replace(txtNovo.Trim(), @"\s+", " ");
        }

        private string removeTagScript(string texto)
        {
            int indexScript = 0;

            while (true)
            {
                indexScript = texto.IndexOf("<script");

                if (indexScript > 0)
                    texto = texto.Remove(indexScript, (texto.IndexOf("</script>") + 9) - indexScript);
                else
                    break;
            }

            return texto;
        }
    }
}