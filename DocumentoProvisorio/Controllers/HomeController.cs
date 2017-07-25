using Busca_LegalDAO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Dynamic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DocumentoProvisorio.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(string action)
        {
            string strPath = Path.Combine(Server.MapPath("~/novos/"));

            HttpPostedFileBase arquivo = Request.Files[0];

            if (arquivo.ContentLength > 0)
            {
                if (!Directory.Exists(strPath))
                    Directory.CreateDirectory(strPath);

                string savedFileName = Path.Combine(strPath, Path.GetFileName(arquivo.FileName.Substring(arquivo.FileName.LastIndexOf("\\") + 1)));
                arquivo.SaveAs(savedFileName);

                /*Criando o diretorio do .ZIP*/
                strPath = savedFileName.Substring(0, savedFileName.ToLower().IndexOf(".zip"));

                if (System.IO.File.Exists(savedFileName))
                {
                    // Salvar o arquivo
                    try
                    {
                        string fileType = arquivo.ContentType;
                        //string arquivosRejeitados = savedFileName;

                        if (fileType == "application/x-zip-compressed")
                        {
                            using (ZipArchive archive = ZipFile.OpenRead(savedFileName))
                            {
                                foreach (ZipArchiveEntry entry in archive.Entries)
                                {
                                    string entryExtension;
                                    //string entryFileType;

                                    entryExtension = entry.FullName;
                                    //entryFileType = "";

                                    if (entryExtension.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ||
                                            entryExtension.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
                                                entryExtension.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                                    entryExtension.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                                    {

                                        if (!Directory.Exists(strPath))
                                            Directory.CreateDirectory(strPath);

                                        //entryFileType = entryExtension.Split('.')[entryExtension.Split('.').Length - 1];

                                        savedFileName = entryExtension;
                                        entry.ExtractToFile(Path.Combine(strPath, savedFileName));

                                        //FileInfo f = new FileInfo(Path.Combine(strPath, savedFileName));
                                        //tamanhoArquivo = f.Length;

                                        /*Neto - *Inicio* Obter os arquivos recem inseridos e gerar string*/

                                        //fs = new FileStream(Path.Combine(strPath, savedFileName), FileMode.Open, FileAccess.Read);
                                        //fileData = null;
                                        //using (var binaryReader = new BinaryReader(fs))
                                        //{
                                        //fileData = binaryReader.ReadBytes((int)fs.Length);
                                        //}

                                        //stringFromByteArray = System.Text.Encoding.Default.GetString(fileData);

                                        //fs.Close();

                                        /*Neto - **Fim** Obter os arquivos recem inseridos e gerar string*/
                                    }
                                    else
                                    {
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }

            /*Fim - Extraindo arquivos ZIP*/

            /*Isolar em Metodo*/
            List<string> listItensZip = new List<string>();

            var listItensDirectory = Directory.GetFiles(strPath);

            foreach (var item in listItensDirectory)
                listItensZip.Add(item.Substring(item.LastIndexOf("\\") + 1) + "|" + "~" + item.Substring(item.IndexOf("\\novos")).Replace("\\", "/"));

            listItensZip = listItensZip.OrderByDescending(x => x).ToList();
            /*Fim - Isolar em Metodo*/

            ViewBag.ListItens = listItensZip;
            ViewBag.Path = listItensDirectory[0].Substring(0, listItensDirectory[0].LastIndexOf("\\")).Replace(@"\", "|");

            return View("Index");
        }

        /// <summary>
        /// Leitura e gravação na base do arquivo.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ProcessarArquivosDocProvisorio(string path)
        {
            new BuscaLegalDao("web").ProcessarArquivosDocProvisorio(path);

            #region "comentado"
            //OleDbConnection _olecon = null;
            //OleDbCommand _oleCmd = null;

            //path = path.Replace("|", @"\");

            //var listItensDirectory = Directory.GetFiles(path).ToList();

            //string _Arquivo = listItensDirectory.Find(x => x.ToLower().Contains(".xls") || x.ToLower().Contains(".xlsx"));
            //string _StringConexao = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;ReadOnly=True';", _Arquivo);
            //string _Consulta = string.Empty;

            //try
            //{
            //    _olecon = new OleDbConnection(_StringConexao);
            //    _olecon.Open();

            //    _oleCmd = new OleDbCommand();
            //    _oleCmd.Connection = _olecon;
            //    _oleCmd.CommandType = CommandType.Text;

            //    _oleCmd.CommandText = "SELECT * FROM [plan1$]";
            //    OleDbDataReader reader = _oleCmd.ExecuteReader();

            //    List<dynamic> listItens = new List<dynamic>();

            //    dynamic objItem;

            //    while (reader.Read())
            //    {
            //        objItem = new ExpandoObject();

            //        objItem.Uf = reader.GetValue(0);
            //        objItem.Especie = reader.GetString(1);
            //        objItem.Orgao = reader.GetString(2);
            //        objItem.Numero = reader.GetValue(3);
            //        objItem.Data_publicacao = reader.GetValue(4);
            //        objItem.Tipo_documento = reader.GetString(5);
            //        objItem.Pdf = reader.GetString(6);

            //        listItens.Add(objItem);
            //    }

            //    reader.Close();

            //    foreach (dynamic itensList in listItens)
            //    {
            //        string caminhoPdf = path + @"\" + itensList.Pdf;
            //        itensList.conteudoPdf = "Ler Arquivo em PDF aqui";
            //    }
            //}
            //catch (Exception ex)
            //{
            //}
            //finally
            //{
            //    if (_oleCmd != null)
            //        _oleCmd.Dispose();

            //    _oleCmd = null;

            //    if (_olecon != null)
            //    {
            //        if (_olecon.State == ConnectionState.Open)
            //            _olecon.Close();

            //        _olecon.Dispose();
            //    }
            //    _olecon = null;
            //}
            #endregion

            return View("Index");
        }
    }
}