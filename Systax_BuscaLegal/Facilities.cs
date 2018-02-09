using Busca_LegalDAO;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Systax_BuscaLegal
{
    public class Facilities
    {
        public string getHtmlPaginaByGet(string url, string encode, string cookie = "", string language = "")
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

        public string obterDadosColection(string texto, string dado)
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

        public string obterEmentaTexto(string texto)
        {
            string textoRetorno = string.Empty;

            var listaEmenta = Regex.Split(texto.ToLower(), "class=resoluoassunto").ToList();

            listaEmenta.RemoveAt(0);

            listaEmenta.ForEach(x => textoRetorno += "<" + x.Substring(0, x.IndexOf("</p>")));

            return textoRetorno;
        }

        public int obterPontoCorte(string titulo)
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

        public dynamic obterEstruturaProcessarUrl_RfbLeg(string url)
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
            catch (Exception)
            {
            }

            return objUrll;
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

        public string GerarHash(string valor)
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

        public string ObterStringLimpa(string y)
        {
            bool key = false;
            string txtNovo = " ";

            y.ToList().ForEach(delegate(char x) { key = x.ToString().Equals("<") || (key && !x.ToString().Equals(">")); txtNovo += !key && !x.ToString().Equals(">") ? x.ToString() : string.Empty; });

            return Regex.Replace(txtNovo.Trim(), @"\s+", " ");
        }

        public string ObterStringLimpa(string y, string corte)
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

        public string removeTagScript(string texto)
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

        public void ProcessarDocsCEF(List<dynamic> listaUrl)
        {
            string urlTratada = string.Empty;
            List<string> itensNaoMapeados = new List<string>();

            foreach (var nivel1_item in listaUrl)
            {
                foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                {
                    try
                    {
                        dynamic itemListaVez = new ExpandoObject();
                        itemListaVez.ListaEmenta = new List<dynamic>();

                        urlTratada = itemLista_Nivel2.Url;
                        itemListaVez.Url = urlTratada;
                        itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                        dynamic dadosEmenta = new ExpandoObject();
                        dadosEmenta.ListaArquivos = new List<ArquivoUpload>();

                        string nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), (Regex.Replace(urlTratada.Split('|')[1], "[^0-9]+", string.Empty) + ".pdf"));

                        using (WebClient webClient = new WebClient())
                        {
                            webClient.DownloadFile(urlTratada.Substring(0, urlTratada.IndexOf("|")), nomeArq);
                        }

                        string dadosPdf = new Facilities().LeArquivo(nomeArq);

                        byte[] arrayFile = File.ReadAllBytes(nomeArq);

                        dadosEmenta.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = "pdf", NomeArquivo = (Regex.Replace(urlTratada.Split('|')[1], "[^0-9]+", string.Empty) + ".pdf") });

                        /** Outros **/
                        dadosEmenta.DescSigla = string.Empty;
                        dadosEmenta.HasContent = false;

                        /** Arquivo **/
                        dadosEmenta.Tipo = 3;
                        dadosEmenta.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                        string[] listItens = urlTratada.Split('|');

                        string titulo = string.Empty;
                        string numero = string.Empty;

                        titulo = listItens[3] + " " + listItens[1];

                        titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                        /** Default **/
                        dadosEmenta.Publicacao = listItens[2];
                        dadosEmenta.Sigla = "CEF - Caixa Econômica Federal";
                        dadosEmenta.Republicacao = string.Empty;
                        dadosEmenta.Ementa = listItens[4];
                        dadosEmenta.TituloAto = titulo.Substring(titulo.IndexOf(":") + 1);
                        dadosEmenta.Especie = listItens[3].Substring(listItens[3].IndexOf(":") + 1); ;
                        dadosEmenta.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                        dadosEmenta.DataEdicao = string.Empty;
                        dadosEmenta.Texto = dadosPdf;
                        dadosEmenta.Hash = new Facilities().GerarHash(new Facilities().removerCaracterEspecial(dadosPdf));
                        dadosEmenta.IdFila = itemLista_Nivel2.Id;

                        dadosEmenta.Escopo = "FED";
                        dadosEmenta.IdFila = itemLista_Nivel2.Id;

                        itemListaVez.ListaEmenta.Add(dadosEmenta);

                        File.Delete(nomeArq);

                        dynamic itemFonte = new ExpandoObject();

                        itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };
                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });

                        Thread.Sleep(2000);
                    }
                    catch (Exception)
                    {
                        itensNaoMapeados.Add(urlTratada);
                    }
                }
            }

            //string novoNcm = "URL_DOC\n";
            //itensNaoMapeados.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
            //File.WriteAllText(@"C:\Temp\CEF_ERROS.csv", novoNcm);
        }

        public void ProcessarDocsNFE(List<dynamic> listaUrl)
        {
            string urlTratada = string.Empty;
            List<string> itensNaoMapeados = new List<string>();

            foreach (var nivel1_item in listaUrl)
            {
                foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                {
                    try
                    {
                        dynamic itemListaVez = new ExpandoObject();
                        itemListaVez.ListaEmenta = new List<dynamic>();

                        urlTratada = itemLista_Nivel2.Url;
                        itemListaVez.Url = urlTratada;
                        itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                        dynamic dadosEmenta = new ExpandoObject();
                        dadosEmenta.ListaArquivos = new List<ArquivoUpload>();

                        string nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), Guid.NewGuid().ToString() + ".pdf");

                        using (WebClient webClient = new WebClient())
                        {
                            webClient.DownloadFile(urlTratada.Substring(0, urlTratada.IndexOf("|")), nomeArq);
                        }

                        string dadosPdf = new Facilities().LeArquivo(nomeArq);

                        byte[] arrayFile = File.ReadAllBytes(nomeArq);

                        dadosEmenta.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = "pdf", NomeArquivo = nomeArq.Substring(nomeArq.LastIndexOf("\"") + 1) });

                        /** Outros **/
                        dadosEmenta.DescSigla = string.Empty;
                        dadosEmenta.HasContent = false;

                        /** Arquivo **/
                        dadosEmenta.Tipo = 3;
                        dadosEmenta.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                        string[] listItens = urlTratada.Split('|');

                        string titulo = listItens[1];
                        string numero = "0";

                        //titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                        /** Default **/
                        dadosEmenta.Publicacao = string.Empty;
                        dadosEmenta.Sigla = "NFE - Nota Fiscal Eletrônica";
                        dadosEmenta.Republicacao = string.Empty;
                        dadosEmenta.Ementa = listItens[2];
                        dadosEmenta.TituloAto = titulo;
                        dadosEmenta.Especie = "Nota Técnica";
                        dadosEmenta.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                        dadosEmenta.DataEdicao = string.Empty;
                        dadosEmenta.Texto = dadosPdf;
                        dadosEmenta.Hash = new Facilities().GerarHash(new Facilities().removerCaracterEspecial(dadosPdf));
                        dadosEmenta.IdFila = itemLista_Nivel2.Id;

                        dadosEmenta.Escopo = "FED";
                        dadosEmenta.IdFila = itemLista_Nivel2.Id;

                        itemListaVez.ListaEmenta.Add(dadosEmenta);

                        File.Delete(nomeArq);

                        dynamic itemFonte = new ExpandoObject();

                        itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };
                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });

                        Thread.Sleep(2000);
                    }
                    catch (Exception)
                    {
                        itensNaoMapeados.Add(urlTratada);
                    }
                }
            }

            //string novoNcm = "URL_DOC\n";
            //itensNaoMapeados.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
            //File.WriteAllText(@"C:\Temp\NFE_ERROS.csv", novoNcm);
        }

        internal void ProcessarDocsMFZ(List<dynamic> listaUrl)
        {
            string urlTratada = string.Empty;
            List<string> itensNaoMapeados = new List<string>();

            foreach (var nivel1_item in listaUrl)
            {
                foreach (var itemLista_Nivel2 in nivel1_item.Lista_Nivel2)
                {
                    try
                    {
                        dynamic itemListaVez = new ExpandoObject();
                        itemListaVez.ListaEmenta = new List<dynamic>();

                        urlTratada = itemLista_Nivel2.Url;
                        itemListaVez.Url = urlTratada;
                        itemListaVez.IdUrl = itemLista_Nivel2.IdUrl;

                        dynamic dadosEmenta = new ExpandoObject();
                        dadosEmenta.ListaArquivos = new List<ArquivoUpload>();

                        string dadosPdf = string.Empty;
                        string htmlRaiz = getHtmlPaginaByGet(urlTratada.Split('|')[0], string.Empty);

                        htmlRaiz = htmlRaiz.Replace("\"", string.Empty).Replace("\n", " ").Replace("\r", " ").Replace("\t", " ");

                        /*Captura CSS*/

                        var listaCss = Regex.Split(htmlRaiz, "<link").ToList();

                        listaCss.RemoveAt(0);

                        List<string> listaCssTratada = new List<string>();

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

                        if (htmlRaiz.Contains("<style"))
                            descStyle = htmlRaiz.Substring(htmlRaiz.IndexOf("<style"))
                                                       .Substring(0, htmlRaiz.Substring(htmlRaiz.IndexOf("<style")).IndexOf("</style>") + 8);

                        /**Fim Captura CSS**/
                        novaCss += " " + descStyle;

                        var listPdf = Regex.Split(htmlRaiz, "<a").ToList();

                        listPdf.RemoveAll(x => !x.Contains(".pdf"));

                        if (listPdf.Count >= 1)
                        {
                            try
                            {
                                string nomeArq = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory.ToString(), Guid.NewGuid().ToString() + ".pdf");

                                using (WebClient webClient = new WebClient())
                                {
                                    webClient.DownloadFile(listPdf[0].Substring(listPdf[0].IndexOf("href=") + 5).Substring(0, listPdf[0].Substring(listPdf[0].IndexOf("href=") + 5).IndexOf(" ")), nomeArq);
                                }

                                byte[] arrayFile = File.ReadAllBytes(nomeArq);

                                dadosEmenta.ListaArquivos.Add(new ArquivoUpload() { conteudoArquivo = arrayFile, ExtensaoArquivo = "pdf", NomeArquivo = removerCaracterEspecial(urlTratada.Split('|')[1].Replace(" ", string.Empty)) });

                                File.Delete(nomeArq);
                            }
                            catch (Exception)
                            {
                            }
                        }

                        /** Outros **/
                        dadosEmenta.DescSigla = string.Empty;
                        dadosEmenta.HasContent = false;

                        /** Arquivo **/
                        dadosEmenta.Tipo = 3;
                        dadosEmenta.Metadado = "{" + string.Format("\"UF\":\"{0}\"", "FED") + "}";

                        string[] listItens = urlTratada.Split('|');

                        dadosPdf = htmlRaiz.Substring(htmlRaiz.IndexOf("<h1 class=documentFirstHeading")).Substring(0, htmlRaiz.Substring(htmlRaiz.IndexOf("<h1 class=documentFirstHeading")).LastIndexOf("União") + 5) + "</p>";

                        string publicacao = dadosPdf.Substring(dadosPdf.IndexOf("<span class=documentPublished")).Substring(0, dadosPdf.Substring(dadosPdf.IndexOf("<span class=documentPublished")).IndexOf(","));

                        dadosPdf = dadosPdf.Remove(dadosPdf.IndexOf("<div id=viewlet-social-like"), (dadosPdf.IndexOf("<span class=documentPublished") - dadosPdf.IndexOf("<div id=viewlet-social-like")));

                        string titulo = listItens[1];
                        string numero = string.Empty;

                        titulo.Replace(".", string.Empty).ToList().ForEach(delegate(char x) { numero += char.IsNumber(x) && (numero.Equals(string.Empty) || !numero.Contains(".")) ? x.ToString() : numero.Equals(string.Empty) ? string.Empty : "."; });

                        /** Default **/
                        dadosEmenta.Publicacao = ObterStringLimpa(publicacao);
                        dadosEmenta.Sigla = "MF - Ministério da Fazenda";
                        dadosEmenta.Republicacao = string.Empty;
                        dadosEmenta.Ementa = listItens[1];
                        dadosEmenta.TituloAto = titulo;
                        dadosEmenta.Especie = listItens[2].Substring(0, listItens[2].IndexOf(" "));
                        dadosEmenta.NumeroAto = Regex.Replace(numero, "[^0-9]+", string.Empty);
                        dadosEmenta.DataEdicao = string.Empty;
                        dadosEmenta.Texto = novaCss + dadosPdf;
                        dadosEmenta.Hash = GerarHash(removerCaracterEspecial(dadosPdf));
                        dadosEmenta.IdFila = itemLista_Nivel2.Id;

                        dadosEmenta.Escopo = "FED";
                        dadosEmenta.IdFila = itemLista_Nivel2.Id;

                        itemListaVez.ListaEmenta.Add(dadosEmenta);

                        dynamic itemFonte = new ExpandoObject();

                        itemFonte.Lista_Nivel2 = new List<dynamic>() { itemListaVez };
                        new BuscaLegalDao().AtualizarFontes(new List<dynamic>() { itemFonte });

                        Thread.Sleep(2000);
                    }
                    catch (Exception)
                    {
                        itensNaoMapeados.Add(urlTratada);
                    }
                }
            }

            //string novoNcm = "URL_DOC\n";
            //itensNaoMapeados.ForEach(x => novoNcm += x.Replace("\n", "|") + "\n");
            //File.WriteAllText(@"C:\Temp\MFZ_ERROS.csv", novoNcm);
        }

        public static int GetEmbVersion()
        {
            int ieVer = GetBrowserVersion();

            if (ieVer > 9)
                return ieVer * 1000 + 1;

            if (ieVer > 7)
                return ieVer * 1111;

            return 7000;
        } // End Function GetEmbVersion

        public static void FixBrowserVersion()
        {
            string appName = System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetExecutingAssembly().Location);
            FixBrowserVersion(appName);
        }

        public static void FixBrowserVersion(string appName)
        {
            FixBrowserVersion(appName, GetEmbVersion());
        } // End Sub FixBrowserVersion

        // FixBrowserVersion("<YourAppName>", 9000);
        public static void FixBrowserVersion(string appName, int ieVer)
        {
            FixBrowserVersion_Internal("HKEY_LOCAL_MACHINE", appName + ".exe", ieVer);
            FixBrowserVersion_Internal("HKEY_CURRENT_USER", appName + ".exe", ieVer);
            FixBrowserVersion_Internal("HKEY_LOCAL_MACHINE", appName + ".vshost.exe", ieVer);
            FixBrowserVersion_Internal("HKEY_CURRENT_USER", appName + ".vshost.exe", ieVer);
        } // End Sub FixBrowserVersion 

        private static void FixBrowserVersion_Internal(string root, string appName, int ieVer)
        {
            try
            {
                //For 64 bit Machine 
                if (Environment.Is64BitOperatingSystem)
                    Microsoft.Win32.Registry.SetValue(root + @"\Software\Wow6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", appName, ieVer);
                else  //For 32 bit Machine 
                    Microsoft.Win32.Registry.SetValue(root + @"\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", appName, ieVer);


            }
            catch (Exception)
            {
                // some config will hit access rights exceptions
                // this is why we try with both LOCAL_MACHINE and CURRENT_USER
            }
        } // End Sub FixBrowserVersion_Internal 

        public static int GetBrowserVersion()
        {
            // string strKeyPath = @"HKLM\SOFTWARE\Microsoft\Internet Explorer";
            string strKeyPath = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer";
            string[] ls = new string[] { "svcVersion", "svcUpdateVersion", "Version", "W2kVersion" };

            int maxVer = 0;
            for (int i = 0; i < ls.Length; ++i)
            {
                object objVal = Microsoft.Win32.Registry.GetValue(strKeyPath, ls[i], "0");
                string strVal = System.Convert.ToString(objVal);
                if (strVal != null)
                {
                    int iPos = strVal.IndexOf('.');
                    if (iPos > 0)
                        strVal = strVal.Substring(0, iPos);

                    int res = 0;
                    if (int.TryParse(strVal, out res))
                        maxVer = Math.Max(maxVer, res);
                } // End if (strVal != null)

            } // Next i

            return maxVer;
        } // End Function GetBrowserVersion 

        public void CorrecaoData()
        {
            var listDocs = new BuscaLegalDao().ObterDocsCorrecao();

            List<dynamic> listFinal = new List<dynamic>();
            List<string> listaTratamento2 = new List<string>();
            List<string> listaErros = new List<string>();

            string dataReformulada = string.Empty;

            foreach (var item in listDocs)
            {
                dynamic itemFinal = new ExpandoObject();

                itemFinal.Id = item.Id;

                string itemCorrigir = item.Texto;

                dataReformulada = string.Empty;

                itemCorrigir.Replace(" ", "|").Replace("º", string.Empty).Replace("-", "/").ToList().ForEach(delegate(char itemTratar)
                {
                    #region "ID fonte 1,2,7,8,9,10,11,13,14,15"

                    try
                    {
                        if (char.IsNumber(itemTratar) && string.IsNullOrEmpty(dataReformulada))
                            dataReformulada = itemTratar.ToString();

                        else if ((char.IsNumber(itemTratar) || "/.".Contains(itemTratar.ToString())) && !string.IsNullOrEmpty(dataReformulada) && !dataReformulada.Contains("|"))
                            dataReformulada += itemTratar.ToString();

                        else if (!dataReformulada.Contains("|") && !string.IsNullOrEmpty(dataReformulada))
                        {
                            dataReformulada += "|";

                            if (dataReformulada.Length <= 7)
                                dataReformulada = string.Empty;
                        }
                    }
                    catch (Exception)
                    {
                    }

                    #endregion
                });

                dataReformulada = dataReformulada.Replace("|", string.Empty).Replace(".", "/");

                try
                {
                    itemFinal.DataFinal = TratamentoDateTime(dataReformulada).ToString("yyyy-MM-dd");
                    listFinal.Add(itemFinal);
                }
                catch (Exception)
                {
                    listaErros.Add(item.Id + ";" + item.Texto);
                    listaTratamento2.Add(itemCorrigir);
                }
            }

            #region "CORREÇÂO NÍVEL 2"

            List<string> listaMeses = new List<string>() { "janeiro|01", "fevereiro|02", "março|03", "marco|03", "abril|04", "maio|05", "junho|06", "julho|07", "agosto|08", "setembro|09", "outubro|10", "novembro|11", "dezembro|12" };

            listaTratamento2.ForEach(delegate(string itemCorrNivel2)
            {
                var listItensValidar = Regex.Split(itemCorrNivel2, "de ");

                if (!string.IsNullOrEmpty(itemCorrNivel2.Trim()))
                {
                    for (int i = 0; i < listItensValidar.Length; i++)
                    {
                        try
                        {
                            if (listaMeses.Exists(x => x.Split('|')[0].Contains(listItensValidar[i].Trim())))
                            {
                                dataReformulada = Regex.Replace(listItensValidar[i - 1], "[^0-9]+", string.Empty) + "/" + listaMeses.Find(x => x.Split('|')[0].Contains(listItensValidar[i].Trim())).Split('|')[1] + "/" + Regex.Replace(listItensValidar[i + 1], "[^0-9]+", string.Empty);

                                dynamic itemFinal = new ExpandoObject();

                                itemFinal.Id = int.Parse(listaErros.Find(x => x.Contains(itemCorrNivel2.Trim())).Split(';')[0]);
                                itemFinal.DataFinal = TratamentoDateTime(dataReformulada).ToString("yyyy-MM-dd");

                                listaErros.RemoveAll(x => x.Contains(itemCorrNivel2.Trim()));
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            });

            #endregion

            if (listaErros.Count > 0)
            {
                string novoNcm = "ID;data\n";
                listaErros.ForEach(x => novoNcm += x + "\n");
                File.WriteAllText(@"C:\Temp\ConversaoDateTime.csv", novoNcm);
            }

            new BuscaLegalDao().AtualizarDocsCorrecao(listFinal);
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

        public void GravaArquivoLogTxtContinuo(string status, string pathFile = "", string textoLog = "")
        {
            string caminhoArq = pathFile.Trim().Equals(string.Empty) ? @"c:\temp\logExecao.txt" : pathFile;
            string log = (textoLog.Trim().Equals(string.Empty) ? "Última Execução do serviço: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " - " : textoLog) + status + Environment.NewLine;

            if (!System.IO.File.Exists(caminhoArq))
                System.IO.File.Create(caminhoArq).Dispose();

            System.IO.StreamWriter file = new System.IO.StreamWriter(caminhoArq, true);

            file.WriteLine(log);
            file.Close();
            file.Dispose();
        }

        public void GravaArquivoLogCsv(List<string> listItens, string tituloConteudo, string nomeArquivo)
        {
            if (listItens.Count > 0)
            {
                tituloConteudo = tituloConteudo.Contains("\n") ? tituloConteudo : tituloConteudo + "\n";
                nomeArquivo = nomeArquivo.ToLower().Contains(".csv") ? nomeArquivo : nomeArquivo + ".csv";

                listItens.ForEach(x => tituloConteudo += x.Replace("\n", "|") + "\n");
                File.WriteAllText(@"C:\Temp\" + nomeArquivo, tituloConteudo);
            }
        }
    }
}