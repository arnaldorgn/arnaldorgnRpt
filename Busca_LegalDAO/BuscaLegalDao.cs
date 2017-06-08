using Npgsql;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Busca_LegalDAO
{
    public class BuscaLegalDao
    {
        NpgsqlConnection objConection;
        NpgsqlCommand objComando;

        string tipoProcessamento;

        public BuscaLegalDao()
        {
            this.objConection = new NpgsqlConnection(System.Configuration.ConfigurationSettings.AppSettings["strConn"].ToString());
            this.objComando = new NpgsqlCommand();

            this.tipoProcessamento = System.Configuration.ConfigurationSettings.AppSettings["modProc"].ToString();
        }

        #region "Métodos Principais"

        public void AtualizarFontes(List<dynamic> objDados)
        {
            try
            {
                this.objComando.Connection = this.objConection;

                if (this.tipoProcessamento.Equals("u"))
                {
                    this.objConection.Open();

                    int idFonte = 0;

                    foreach (var itemNivel1 in objDados)
                    {
                        //this.ObterParametrosProcedure(itemNivel1);

                        this.objComando.Parameters.Clear();

                        this.objComando.Parameters.Add("pNome", itemNivel1.Indexacao.ToString());
                        this.objComando.Parameters.Add("pPeriodicidade", "1");
                        this.objComando.Parameters.Add("pUltLeitura", DateTime.Now.ToString("yyyy-MM-dd"));
                        this.objComando.Parameters.Add("pRelevancia", 1);

                        this.objComando.CommandTimeout = 0;
                        this.objComando.CommandType = System.Data.CommandType.StoredProcedure;
                        this.objComando.CommandText = "fc_busLegal_inserir_fonte";

                        idFonte = (int)this.objComando.ExecuteScalar();

                        itemNivel1.Id = idFonte;
                    }

                    this.objConection.Close();
                }

                this.AtualizarUrls(objDados);
            }
            catch (Exception ex)
            {
                ex.Source = "Estagio2|Nivel1";
                InserirLogErro(ex, string.Empty, string.Empty);
            }
            finally
            {
                this.objComando.Connection.Dispose();
            }
        }

        protected void AtualizarUrls(List<dynamic> objDados)
        {
            try
            {
                this.objConection.Open();

                int idUrl = 0;

                foreach (var itemNivel1 in objDados)
                {
                    /*Inserindo URL nivel 1*/
                    /*Todo: Incluir fila para inclusão das URL*/
                    if (this.tipoProcessamento.Equals("u"))
                    {
                        this.objComando.Parameters.Clear();

                        this.objComando.Parameters.Add("pIdFonte", itemNivel1.Id);
                        this.objComando.Parameters.Add("pUrl", itemNivel1.Url);
                        this.objComando.Parameters.Add("pNivelProfundiade", 1);
                        this.objComando.Parameters.Add("pIsDoc", DBNull.Value);
                        this.objComando.Parameters.Add("pDtAtualizacao", DateTime.Now.ToString("yyyy-MM-dd"));
                        this.objComando.Parameters.Add("pIdUrl", DBNull.Value);

                        this.objComando.CommandTimeout = 0;
                        this.objComando.CommandType = System.Data.CommandType.StoredProcedure;
                        this.objComando.CommandText = "fc_busLegal_inserir_urls";

                        idUrl = (int)this.objComando.ExecuteScalar();
                    }

                    /*Inserindo URL nivel 2*/
                    /*Todo: Incluir fila para inclusão das URL*/
                    foreach (var itemNivel2 in itemNivel1.Lista_Nivel2)
                    {
                        if (this.tipoProcessamento.Equals("u"))
                        {
                            this.objComando.Parameters.Clear();

                            this.objComando.Parameters.Add("pIdFonte", itemNivel1.Id);
                            this.objComando.Parameters.Add("pUrl", itemNivel2.Url);
                            this.objComando.Parameters.Add("pNivelProfundiade", 2);
                            this.objComando.Parameters.Add("pIsDoc", "1");
                            this.objComando.Parameters.Add("pDtAtualizacao", DateTime.Now.ToString("yyyy-MM-dd"));
                            this.objComando.Parameters.Add("pIdUrl", idUrl);

                            this.objComando.CommandTimeout = 0;
                            this.objComando.CommandType = System.Data.CommandType.StoredProcedure;
                            this.objComando.CommandText = "fc_busLegal_inserir_urls";

                            var result = (int)this.objComando.ExecuteScalar();

                            itemNivel2.IdUrl = result;
                        }

                        if (this.tipoProcessamento.Equals("d"))
                            this.AtualizarDocs(itemNivel2);
                    }
                }
            }
            catch (Exception ex)
            {
                ex.Source = "Estagio2|Nivel2";
                InserirLogErro(ex, string.Empty, string.Empty);
            }
        }

        protected void AtualizarDocs(dynamic dadosEmenta)
        {
            try
            {
                var objListInsert = ((List<dynamic>)dadosEmenta.ListaEmenta);

                int idPai = 0;

                foreach (var itemEmenta in objListInsert.OrderBy(x => x.Tipo))
                {

                    try
                    {
                        this.objComando.Parameters.Clear();

                        this.objComando.Parameters.Add("pTitulo", objListInsert.Find(x => x.Tipo == 3).TituloAto);
                        this.objComando.Parameters.Add("pDataPublicacao", objListInsert.Find(x => x.Tipo == 3).Publicacao);
                        this.objComando.Parameters.Add("pdatarepublicacao", objListInsert.Find(x => x.Tipo == 3).Republicacao);
                        this.objComando.Parameters.Add("pEmenta", objListInsert.Find(x => x.Tipo == 3).Ementa);
                        this.objComando.Parameters.Add("pTexto", itemEmenta.Texto);
                        this.objComando.Parameters.Add("pFlagVisao", itemEmenta.Tipo.ToString());
                        this.objComando.Parameters.Add("pIdPai", idPai);
                        this.objComando.Parameters.Add("pIdUrl", dadosEmenta.IdUrl);
                        this.objComando.Parameters.Add("pEspecie", objListInsert.Find(x => x.Tipo == 3).Especie);
                        this.objComando.Parameters.Add("pUf", objListInsert.Find(x => x.Tipo == 3).Escopo);
                        this.objComando.Parameters.Add("pDtAtualizacao", DateTime.Now.ToString("yyyy-MM-dd"));
                        this.objComando.Parameters.Add("pMetadado", objListInsert.Find(x => x.Tipo == 3).Metadado);
                        this.objComando.Parameters.Add("pHash", itemEmenta.Hash);

                        this.objComando.Parameters.Add("pSiglaOrgao", !objListInsert.Find(x => x.Tipo == 3).Sigla.Equals(string.Empty) ? string.Format("{0} - {1}", objListInsert.Find(x => x.Tipo == 3).Sigla, objListInsert.Find(x => x.Tipo == 3).DescSigla) : DBNull.Value);
                        this.objComando.Parameters.Add("pNumeroAto", objListInsert.Find(x => x.Tipo == 3).NumeroAto);
                        this.objComando.Parameters.Add("pDataEdicao", objListInsert.Find(x => x.Tipo == 3).DataEdicao);
                        this.objComando.Parameters.Add("pIdFila", itemEmenta.IdFila);

                        if (itemEmenta.HasContent)
                        {
                            this.objComando.Parameters.Add("pByteArrPdf", itemEmenta.ByteArrPdf);
                            this.objComando.Parameters.Add("pNomeArquivo", itemEmenta.NomeArquivo);
                            this.objComando.Parameters.Add("pExtensao", itemEmenta.ExtensaoArquivo);
                        }
                        else
                        {
                            this.objComando.Parameters.Add("pByteArrPdf", DBNull.Value);
                            this.objComando.Parameters.Add("pNomeArquivo", DBNull.Value);
                            this.objComando.Parameters.Add("pExtensao", DBNull.Value);
                        }

                        this.objComando.CommandTimeout = 0;
                        this.objComando.CommandType = System.Data.CommandType.StoredProcedure;
                        this.objComando.CommandText = "fc_busLegal_inserir_docs";

                        var result = (int)this.objComando.ExecuteScalar();

                        idPai = idPai == 0 ? result : idPai;
                    }
                    catch (Exception ex)
                    {
                        ex.Source = "Estagio2|Nivel3";
                        InserirLogErro(ex, string.Empty, string.Empty);
                    }
                }

                /*Salvando Arquivos que venham a existir no documento*/
                if (objListInsert.Find(x => x.Tipo == 3).ListaArquivos.Count > 0)
                {
                    foreach (ArquivoUpload itemArq in objListInsert.Find(x => x.Tipo == 3).ListaArquivos)
                    {
                        this.objComando.Parameters.Clear();

                        this.objComando.Parameters.Add("pIdConteudo", idPai);
                        this.objComando.Parameters.Add("pConteudo", itemArq.conteudoArquivo);
                        this.objComando.Parameters.Add("pExtensao", itemArq.ExtensaoArquivo);
                        this.objComando.Parameters.Add("pNomeArq", itemArq.NomeArquivo);

                        this.objComando.CommandTimeout = 0;
                        this.objComando.CommandType = System.Data.CommandType.StoredProcedure;
                        this.objComando.CommandText = "fc_busLegal_inserir_docs_lote";

                        this.objComando.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                ex.Source = "Estagio2|Nivel3";
                InserirLogErro(ex, string.Empty, string.Empty);
            }
        }

        public List<dynamic> ObterUrlsParaProcessamento(string tipoFonte)
        {
            /*Remonta a estrutura do objeto para o processamento dos docs*/
            List<dynamic> listaRetorno = new List<dynamic>();

            dynamic itemNivel1 = new ExpandoObject();
            dynamic itemListaDocs;

            try
            {
                this.objComando.Connection = this.objConection;

                this.objComando.Parameters.AddWithValue("pTipoFonte", tipoFonte);

                this.objComando.CommandTimeout = 0;
                this.objComando.CommandType = System.Data.CommandType.StoredProcedure;
                this.objComando.CommandText = "fc_obter_lista_urls_processar_docs";

                this.objConection.Open();

                using (NpgsqlDataReader dataReaderObj = this.objComando.ExecuteReader())
                {
                    while (dataReaderObj.Read())
                    {
                        itemListaDocs = new ExpandoObject();

                        itemListaDocs.IdUrl = int.Parse(dataReaderObj["id_url"].ToString());
                        itemListaDocs.Url = dataReaderObj["url"].ToString();
                        itemListaDocs.Id = dataReaderObj["id"].ToString();

                        listaRetorno.Add(itemListaDocs);
                    }
                }

                itemNivel1.Lista_Nivel2 = listaRetorno;
            }
            catch (Exception ex)
            {
                ex.Source = "Estagio2|ObterUrlsParaProcessamento";
                InserirLogErro(ex, string.Empty, string.Empty);
            }
            finally
            {
                this.objConection.Dispose();
            }

            return new List<dynamic>() { itemNivel1 };
        }

        public void InserirLogErro(Exception ex, string urls, string detalhes)
        {
            try
            {
                this.objComando.Connection = this.objConection;

                this.objComando.Parameters.AddWithValue("pIdUrl", DBNull.Value);
                this.objComando.Parameters.AddWithValue("pIdConteudo", DBNull.Value);
                this.objComando.Parameters.AddWithValue("pStatus", "E");
                this.objComando.Parameters.AddWithValue("pDescricao", string.Format("{0}|{1}|{2}|{3}", urls, ex.Message, ex.Source, detalhes));

                this.objComando.CommandTimeout = 0;
                this.objComando.CommandType = System.Data.CommandType.StoredProcedure;
                this.objComando.CommandText = "fc_busLegal_inserir_erro";

                this.objConection.Open();

                this.objComando.ExecuteNonQuery();
            }
            catch (Exception ex1)
            {
                throw ex1;
            }
            finally
            {
                this.objConection.Dispose();
            }
        }

        #endregion

        #region "Métodos Aux"

        /// <summary>
        /// Obter as propriedades do objFiltro e criar parametros apartir deles
        /// </summary>
        /// <param name="objInserir">obj com os filtros</param>
        private void ObterParametrosProcedure(dynamic objInserir)
        {
            foreach (PropertyInfo propInfo in objInserir.GetType().GetProperties())
                this.objComando.Parameters.Add("p" + propInfo.Name, !string.IsNullOrEmpty(Convert.ToString(propInfo.GetValue(objInserir, null))) ? (propInfo.GetValue(objInserir, null)).GetType() == typeof(List<string>) ? ((List<string>)(propInfo.GetValue(objInserir, null))).ToArray() : propInfo.GetValue(objInserir, null) : DBNull.Value);
        }

        #endregion

        public List<dynamic> ObterDocsCorrecao()
        {
            List<dynamic> listaRetorno = new List<dynamic>();

            try
            {
                dynamic itemListaDocs;


                this.objComando.Connection = this.objConection;

                this.objComando.CommandTimeout = 0;
                this.objComando.CommandType = System.Data.CommandType.Text;
                this.objComando.CommandText = "select id, data_publicacao as texto from conteudo where id_url in (select id from url where id_fonte = 1) and trim(data_publicacao) <> '' and data_publicacao is not null and dt_publicacao is null";

                this.objConection.Open();

                using (NpgsqlDataReader dataReaderObj = this.objComando.ExecuteReader())
                {
                    while (dataReaderObj.Read())
                    {
                        itemListaDocs = new ExpandoObject();

                        itemListaDocs.Id = (int)dataReaderObj["id"];
                        itemListaDocs.Texto = dataReaderObj["texto"].ToString();

                        listaRetorno.Add(itemListaDocs);
                    }
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                this.objConection.Dispose();
            }

            return listaRetorno;
        }

        public void AtualizarDocsCorrecao(List<dynamic> listUpdate)
        {
            try
            {
                this.objComando.Connection = this.objConection;
                this.objConection.Open();

                foreach (var item in listUpdate)
                {
                    this.objComando.Parameters.Clear();

                    this.objComando.Parameters.AddWithValue("@data", item.DataFinal);
                    this.objComando.Parameters.AddWithValue("@id", item.Id);

                    this.objComando.CommandTimeout = 0;
                    this.objComando.CommandType = System.Data.CommandType.Text;
                    this.objComando.CommandText = "update conteudo set dt_publicacao = @data where id = @id";

                    this.objComando.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                this.objConection.Dispose();
            }
        }
    }

    public class ArquivoUpload
    {
        public byte[] conteudoArquivo;

        public string NomeArquivo;

        public string ExtensaoArquivo;
    }
}