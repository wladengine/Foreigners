using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Data;

using EducServLib;
using BDClassLib;

namespace Priem
{
    public class DocsClass
    {
        private DBPriem _bdcInet;
        private DBPriem _bdcInetFiles;
        private string _personId;
        private string _abitId;
        private string _commitId;

        public DocsClass(int personBarcode, int? abitCommitBarcode)
        {
            _bdcInet = new DBPriem();
            _bdcInetFiles = new DBPriem();
            try
            {
                _bdcInet.OpenDatabase(MainClass.connStringOnline);
                _bdcInetFiles.OpenDatabase(MainClass.connStringOnlineFiles);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);                
            }
                        
            _personId = _bdcInet.GetStringValue("SELECT Person.Id FROM Person WHERE Person.Barcode = " + personBarcode);

            if (abitCommitBarcode == null)
                _abitId = null;
            else
            {
                _abitId = _bdcInet.GetStringValue("SELECT qAbiturient.Id FROM qAbiturient WHERE qAbiturient.CommitNumber = " + abitCommitBarcode);
                _commitId = _bdcInet.GetStringValue("SELECT qAbiturient.CommitId FROM qAbiturient WHERE qAbiturient.CommitNumber = " + abitCommitBarcode);
            }
        }        

        public DBPriem BDCInet
        {
            get { return _bdcInet; }
        }

        public void CloseDB()
        {
            _bdcInet.CloseDataBase();
        }

        public void OpenFile(List<KeyValuePair<string, string>> lstFiles)
        {
            try
            {
                foreach (KeyValuePair<string, string> file in lstFiles)
                {
                    byte[] bt = _bdcInetFiles.ReadFile(string.Format("SELECT FileData FROM extAbitFileNames_All WHERE Id = '{0}'", file.Key));

                    string filename = file.Value.Replace(@"\", "-").Replace(@":", "-");

                    StreamWriter sw = new StreamWriter(MainClass.saveTempFolder + filename);
                    BinaryWriter bw = new BinaryWriter(sw.BaseStream);
                    bw.Write(bt);
                    bw.Flush();
                    bw.Close();
                    Process.Start(MainClass.saveTempFolder + filename);
                }
            }
            catch (System.Exception exc)
            {
                WinFormsServ.Error("Ошибка открытия файла: " + exc.Message);
            }
        }

        public List<KeyValuePair<string, string>> UpdateFiles()
        {
            try
            {            
                if (_personId == null)
                    return null;

                List<KeyValuePair<string, string>> lstFiles = new List<KeyValuePair<string, string>>();

                string query = string.Format("SELECT Id, FileName + ' (' + convert(nvarchar, extAbitFileNames_All.LoadDate, 104) + ' ' + convert(nvarchar, extAbitFileNames_All.LoadDate, 108) + ')' + FileExtention AS FileName FROM extAbitFileNames_All WHERE extAbitFileNames_All.PersonId = '{0}' {1} {2}", _personId, 
                    !string.IsNullOrEmpty(_abitId) ? " AND (extAbitFileNames_All.ApplicationId = '" + _abitId + "' OR extAbitFileNames_All.ApplicationId IS NULL)" : "",
                    !string.IsNullOrEmpty(_commitId) ? " AND (extAbitFileNames_All.CommitId = '" + _commitId + "' OR extAbitFileNames_All.CommitId IS NULL)" : "");
                DataSet ds = _bdcInetFiles.GetDataSet(query + " ORDER BY extAbitFileNames_All.LoadDate DESC");
                foreach (DataRow dRow in ds.Tables[0].Rows)
                {
                    lstFiles.Add(new KeyValuePair<string, string>(dRow["Id"].ToString(), dRow["FileName"].ToString()));
                }
                
                return lstFiles;
            }
            catch (System.Exception exc)
            {
                WinFormsServ.Error("Ошибка обновления данных о приложениях: " + exc.Message);
                return null;
            }
        }
    }
}
