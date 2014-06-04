using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

using BDClassLib;
using EducServLib;

namespace Priem
{
    public partial class Migrator : Form
    {
        private DBPriem _bdc;
        private OleDbClass _odc;
        private string _emptyBase = string.Format(@"{0}\Templates\EMPTYAbiturientDB.mdb", Application.StartupPath);
        private string _metroBase = string.Format(@"{0}\Templates\MetroDB.mdb", Application.StartupPath);

        private long _NewId = 1000001;
        private SortedList<string, long> _slIds;
        private SortedList<string, string> slProf = null;
        private SortedList<string, string> slSpec = null;
        ArrayList _alQueries;
       
        //конструктор
        public Migrator()
        {
            InitializeComponent();
            InitItems();
        }

        //дополнительная инициализация
        private void InitItems()
        {
            this.CenterToParent();
            this.MdiParent = MainClass.mainform;

            _bdc = MainClass.Bdc;
            
            if(MainClass.IsPasha())
                ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.SP_Faculty"), false, true);               
            else
                ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.SP_Faculty"), false, false);           

            btnStart.Enabled = false;
            btnMetro.Enabled = false;
        }

        //старт
        private void btnStart_Click(object sender, EventArgs e)
        {
            string newfile = folderBrowser.SelectedPath + "/AbiturientDB.mdb";

            FileInfo fi = new FileInfo(_emptyBase);
            fi.CopyTo(newfile, true);

            _odc = new OleDbClass();
            _odc.OpenDatabase(newfile);
                        
            _alQueries = new ArrayList();
            PrepareRegion();
            MigrateProfSpez();
            MigrateOrders();
            MigrateAbits();
            _odc.ExecuteWithTrasaction(_alQueries);
            MessageBox.Show("Готово!");
            _odc.CloseDataBase();
        }

        public string FacultyId
        {
            get { return ComboServ.GetComboId(cbFaculty); }            
        }

        Dictionary<string, string> _dRegion;
        private void PrepareRegion()
        {
            _dRegion = new Dictionary<string, string>();

            DataSet ds = _bdc.GetDataSet("SELECT * FROM ed.Region");

            foreach (DataRow row in ds.Tables[0].Rows)
                _dRegion.Add(row["Name"].ToString(), row["Id"].ToString());
        }

        //путь сохранения
        private void btnFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbFolder.Text = folderBrowser.SelectedPath;
                btnStart.Enabled = true;
                if(MainClass.IsPasha())
                    btnMetro.Enabled = true;
            }
        }
        
        //миграция
        private void MigrateProfSpez()
        {
            slProf = new SortedList<string, string>();
            slSpec = new SortedList<string, string>();

            string query = @"SELECT DISTINCT LicenseProgramId, FacultyId, LicenseProgramCode, LicenseProgramName  
                       FROM ed.qEntry " + (string.IsNullOrEmpty(FacultyId) ? "" : " WHERE ed.qEntry.FacultyId = " + FacultyId);

            DataSet dsProf = _bdc.GetDataSet(query);

            int newProfId = 1;
            foreach (DataRow rowProf in dsProf.Tables[0].Rows)
            {
                string queryIns = string.Format("INSERT INTO Profession (Id, Name, Code, FacultyId) VALUES ({0}, '{1}', '{2}', {3})", newProfId, rowProf["LicenseProgramName"].ToString(), rowProf["LicenseProgramCode"].ToString(), rowProf["FacultyId"].ToString());
                _odc.ExecuteQuery(queryIns);

                string key = rowProf["FacultyId"].ToString() + "_" + rowProf["LicenseProgramId"].ToString();
                string value = newProfId.ToString();
                slProf.Add(key, value);

                newProfId++;
            }

            query = @"SELECT DISTINCT LicenseProgramId, FacultyId, ProfileId, ProfileName 
                       FROM ed.qEntry WHERE NOT ProfileId IS NULL " + (string.IsNullOrEmpty(FacultyId) ? "" : " AND ed.qEntry.FacultyId = " + FacultyId);

            DataSet dsSpec = _bdc.GetDataSet(query);

            int newSpecId = 1;
            foreach (DataRow rowSpec in dsSpec.Tables[0].Rows)
            {               
                string profId = slProf[rowSpec["FacultyId"].ToString() + "_" + rowSpec["LicenseProgramId"].ToString()];

                string queryIns = string.Format("INSERT INTO Specialization (Id, Name, ProfessionId) VALUES ({0}, '{1}', {2})", newSpecId, rowSpec["ProfileName"].ToString(), profId);
                _odc.ExecuteQuery(queryIns);

                string key = string.Format("{0}_{1}_{2}", rowSpec["FacultyId"], rowSpec["LicenseProgramId"], rowSpec["ProfileId"]);
                string value = newSpecId.ToString();
                slSpec.Add(key, value);

                newSpecId++;
            }
        }
       
        private void MigrateOrders()
        {
            string query = @"SELECT ed.OrderNumbers.*, ed.Protocol.FacultyId, ed.Protocol.StudyFormId, ed.Protocol.StudyBasisId FROM ed.OrderNumbers 
                             INNER JOIN ed.Protocol On ed.OrderNumbers.ProtocolId = ed.Protocol.Id WHERE 1=1 " + GetFilter("ed.Protocol");
            string queryAbits;

            DataSet ds = _bdc.GetDataSet(query);
            
            _slIds = new SortedList<string, long>();

            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                string s = string.Empty;
                if (dr["OrderNum"].ToString().Length != 0)
                {
                    s = string.Format("INSERT INTO Protocol (Id,FacultyId, SectionId, StudyFormId,Name,FromDate) VALUES ({0},{1},{2},{3},'{4}','{5}')", _NewId, dr["FacultyId"].ToString(), dr["StudyFormId"].ToString(), dr["StudyBasisId"].ToString(), dr["OrderNum"].ToString(), dr["OrderDate"].ToString());
                    _alQueries.Add(s);

                    queryAbits = string.Format(@"SELECT ed.extEntryView.AbiturientId FROM ed.extEntryView INNER JOIN ed.qAbiturient ON ed.extEntryView.AbiturientId = ed.qAbiturient.Id INNER JOIN ed.Person ON ed.qAbiturient.PersonId = ed.Person.Id WHERE ed.extEntryView.Id = '{0}' 
                                                 AND ed.Person.NationalityId=1  ", dr["ProtocolId"].ToString());
                    foreach (DataRow drr in _bdc.GetDataSet(queryAbits).Tables[0].Rows)
                    {
                        _slIds.Add(drr["AbiturientId"].ToString(), _NewId);
                    }                 
                    _NewId++;          
                }
                if (dr["OrderNumFor"].ToString().Length != 0)
                {
                    s = string.Format("INSERT INTO Protocol (Id,FacultyId, SectionId, StudyFormId,Name,FromDate) VALUES ({0},{1},{2},{3},'{4}','{5}')", _NewId, dr["FacultyId"].ToString(), dr["StudyFormId"].ToString(), dr["StudyBasisId"].ToString(), dr["OrderNumFor"].ToString(), dr["OrderDateFor"].ToString());
                    _alQueries.Add(s);

                    queryAbits = string.Format(@"SELECT ed.extEntryView.AbiturientId FROM ed.extEntryView INNER JOIN ed.qAbiturient ON ed.extEntryView.AbiturientId = ed.qAbiturient.Id INNER JOIN ed.Person ON ed.qAbiturient.PersonId = ed.Person.Id 
                                                 WHERE ed.extEntryView.Id = '{0}' AND ed.Person.NationalityId <> 1 ", dr["ProtocolId"].ToString());
                    foreach (DataRow drr in _bdc.GetDataSet(queryAbits).Tables[0].Rows)
                    {
                        _slIds.Add(drr["AbiturientId"].ToString(), _NewId);
                    } 
                    _NewId++;          
                }  
            }
        }

        //миграция
        private void MigrateAbits()
        {
            string query = string.Format("SELECT DISTINCT ed.extAbit.Id, ed.extAbit.RegNum, ed.Person.Name, ed.Person.SecondName, ed.Person.Surname, " +
                              "ed.Person.BirthDate, ed.Person.BirthPlace, ed.Person.Sex, " +
                              "ed.Person.PassportTypeId, ed.Person.PassportSeries, ed.Person.PassportNumber, ed.Person.PassportAuthor, ed.Person.PassportDate, " +
                              "ed.Person.NationalityId, ed.Person.RegionId, ed.Person.Phone, ed.Person.Mobiles, " +
                              "ed.extAbit.LanguageId, ed.Person.Privileges, " +
                              "ed.Person.Code, Person.City, Person.Street, Person.House, Person.Korpus, Person.Flat, " +
                              "ed.Person.CodeReal, ed.Person.CityReal, ed.Person.StreetReal, ed.Person.HouseReal, ed.Person.KorpusReal, ed.Person.FlatReal, " +
                              "ed.Person.SchoolTypeId, ed.Person.SchoolCity, ed.Person.SchoolName, ed.Person.SchoolNum, ed.Person.SchoolExitYear, " +
                              "ed.Person.AttestatRegion, ed.Person.AttestatSeries, ed.Person.AttestatNum, (case when ed.extAbit.AttDocOrigin = 1 then 'false' else 'true' end) AS CopyAtt, " +
                              "ed.Person.DiplomSeries, ed.Person.DiplomNum, " +
                              "ed.Person.IsExcellent, ed.Country.Name as Nation, " +
                              "ed.extAbit.FacultyId, ed.extAbit.StudyFormId, ed.extAbit.StudyBasisId, " +
                              "ed.extAbit.LicenseProgramId, ed.extAbit.ProfileId, ed.extAbit.CompetitionId, ed.extAbit.StudyNumber, " +
                              "ed.extAbit.ObrazProgramName, ed.extAbit.ObrazProgramCrypt, " +
                              "ed.extAbit.DocDate, ed.extAbit.HostelEduc, ed.extAbit.IsListener, ed.extAbit.StudyPlanNumber, " +
                              "(Case When ed.extAbit.IsSecond = 1 then 1 else (case when ed.extabit.Isreduced = 1 then 2 else (case when ed.extabit.isparallel = 1 then 3 else 0 end) end) end) AS ListenerTypeId, " +
                              "ed.extEntryView.Id AS EntryProtId, ed.Person.DiplomSeries, ed.Person.DiplomNum, ed.Person.HEExitYear " +
                              "FROM ed.extAbit INNER JOIN ed.Person ON ed.extAbit.PersonId = ed.Person.Id " +
                              "INNER JOIN ed.extEntryView ON ed.extEntryView.AbiturientId = ed.extAbit.Id " +
                              "INNER JOIN ed.OrderNumbers ON ed.extEntryView.Id = ed.OrderNumbers.ProtocolId " +
                              "LEFT JOIN ed.Country ON ed.Country.Id=ed.Person.NationalityId " +
                              "WHERE ed.extAbit.Id IN ({1}) {0}", GetFilter("ed.extAbit"), Util.BuildStringWithCollectionWithApps(_slIds.Keys));


            //string query = string.Format("SELECT DISTINCT Abiturient.SectionId,Abiturient.CompetitionId,case when Abiturient.AttestatCopy>0 then 1 else 0 end as AttestatCopy, Abiturient.SchoolTypeId,Abiturient.ExitYear,Abiturient.FileNum, Abiturient.FacultyId, Abiturient.StudyNumber, Abiturient.Surname, Abiturient.Name, Abiturient.Patronymic, Abiturient.BirthDate, Case When Abiturient.Sex>0 then 1 else 0 end as Sex , Abiturient.Privileges, Abiturient.PasswordTypeId, Abiturient.PaswSeries, Abiturient.PaswNumber, Abiturient.PaswDate,Abiturient.PaswAuthor, Abiturient.ZipCode, Abiturient.Adress, Abiturient.LifeAddress,Abiturient.Phone, Abiturient.CitizenId, Abiturient.RegionId,Abiturient.LanguageId, Abiturient.DocDate,Abiturient.AttestatSeries,Abiturient.AttestatNum,Abiturient.SchoolCity, Abiturient.SchoolName, Abiturient.SchoolNum,Abiturient.ProfessionId,Abiturient.SpecializationId,Abiturient.StudyFormId,Abiturient.GroupNum, qAbitProtocols.EntryOrderId  FROM Abiturient INNER JOIN qAbitProtocols ON qAbitProtocols.Id = Abiturient.Id INNER JOIN Protocol ON qAbitProtocols.EntryOrderId = Protocol.Id WHERE NOT qAbitProtocols.EntryOrderId IS NULL AND NOT Protocol.OrderName IS NULL ") + GetFilter("Abiturient");
            //string query = "SELECT Abiturient.SectionId,Abiturient.CompetitionId,case when Abiturient.AttestatCopy>0 then 1 else 0 end as AttestatCopy, Abiturient.SchoolTypeId,Abiturient.ExitYear,Abiturient.FileNum, Abiturient.FacultyId, Abiturient.StudyNumber, Abiturient.Surname, Abiturient.Name, Abiturient.Patronymic, Abiturient.BirthDate, Case When Abiturient.Sex>0 then 1 else 0 end as Sex , Abiturient.Privileges, Abiturient.PasswordTypeId, Abiturient.PaswSeries, Abiturient.PaswNumber, Abiturient.PaswDate,Abiturient.PaswAuthor, Abiturient.ZipCode, Abiturient.Adress, Abiturient.LifeAddress,Abiturient.Phone, Abiturient.CitizenId, Abiturient.RegionId,Abiturient.LanguageId, Abiturient.DocDate,Abiturient.AttestatSeries,Abiturient.AttestatNum,Abiturient.SchoolCity, Abiturient.SchoolName, Abiturient.SchoolNum,Abiturient.ProfessionId,Abiturient.SpecializationId,Abiturient.StudyFormId,Abiturient.GroupNum, qAbitProtocols.EntryOrderId  FROM Abiturient INNER JOIN qAbitProtocols ON qAbitProtocols.Id = Abiturient.Id WHERE 1=1 " + GetFilter("Abiturient");

            DataSet ds = _bdc.GetDataSet(query);
            if (ds.Tables[0].Rows.Count == 0)
                return;

            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                string zc = dr["Code"].ToString().Replace(" ", "");
                if (zc.Length > 10)
                    zc = zc.Substring(0, 10);

                string pa = dr["PassportAuthor"].ToString();
                if (pa.Length > 250)
                    pa = pa.Substring(0, 250);

                string a = dr["Code"].ToString() + ", " + dr["City"].ToString() + ", " + dr["Street"].ToString() + ", д." + dr["House"].ToString() + (dr["Korpus"].ToString().Length == 0 ? "" : " к." + dr["Korpus"].ToString()) + " кв." + dr["Flat"].ToString();
                if (a.Length > 250)
                    a = a.Substring(0, 250);

                string la = dr["CodeReal"].ToString() + ", " + dr["CityReal"].ToString() + ", " + dr["StreetReal"].ToString() + ", д." + dr["HouseReal"].ToString() + (dr["KorpusReal"].ToString().Length == 0 ? "" : " к." + dr["KorpusReal"].ToString()) + " кв." + dr["FlatReal"].ToString();
                if (dr["CodeReal"].ToString().Length == 0 && dr["CityReal"].ToString().Length == 0 && dr["StreetReal"].ToString().Length == 0 && dr["HouseReal"].ToString().Length == 0 && dr["KorpusReal"].ToString().Length == 0 && dr["FlatReal"].ToString().Length == 0)
                    la = "";
                if (la.Length > 250)
                    la = la.Substring(0, 250);

                string ph = dr["Phone"].ToString() + (dr["Mobiles"].ToString().Length == 0 ? "" : "; " + dr["Mobiles"].ToString());
                if (ph.Length > 100)
                    ph = ph.Substring(0, 100);

                string profId = slProf[dr["FacultyId"].ToString() + "_" + dr["LicenseProgramId"].ToString()];

                string specId;
                if (dr["ProfileId"] == null || dr["ProfileId"].ToString() == "0" || dr["ProfileId"].ToString() == "")
                    specId = "0";
                else
                    specId = slSpec[dr["FacultyId"].ToString() + "_" + dr["LicenseProgramId"].ToString() + "_" + dr["ProfileId"].ToString()];

                string educSeries = MainClass.studyLevelGroupId == 1 ? dr["AttestatSeries"].ToString() : dr["DiplomSeries"].ToString();
                string educNum = MainClass.studyLevelGroupId == 1 ? dr["AttestatNum"].ToString() : dr["DiplomNum"].ToString();
                string educYear = MainClass.studyLevelGroupId == 1 ? dr["SchoolExitYear"].ToString() : dr["HEExitYear"].ToString();
                if (string.IsNullOrEmpty(educYear))
                    educYear = "0";

                string s = string.Format(
                    "INSERT INTO Abiturient (" +
                    "[FileNum], [Name], [Patronymic], [Surname], " +
                    "[Privileges], [IsExcellent], [ListenerTypeId], [IsActualListener], " +
                    "[Hostel], [FacultyId], [ProfessionId], [SpecializationId], " +
                    "[StudyFormId], [SectionId], [CompetitionId], " +
                    "[DocDate], [CitizenId], [RegionId], [LanguageId], " +
                    "[AttestatSeries], [AttestatRegion], [AttestatNum], [AttestatCopy], " +
                    "[SchoolName], [SchoolCity], [SchoolNum], [SchoolTypeId], [ExitYear], " +
                    "[Phone], [ZipCode], [Adress], [LifeAddress], " +
                    "[BirthDate], [Sex], " +
                    "[PasswordTypeId], [PaswSeries], [PaswNumber], [PaswDate], [PaswAuthor], " +
                    "[StudyNumber], [EntryOrderId], [EduProgName], [EduProgKod], [StudyPlanNumber])" +
                    "VALUES (" +
                    "'{0}','{1}','{2}','{3}'," +
                    "'{4}','{5}','{6}','{7}', " +
                    "'{8}','{9}','{10}','{11}'," +
                    "'{12}','{13}','{14}'," +
                    "'{15}','{16}','{17}','{18}'," +
                    "'{19}','{20}','{21}','{22}'," +
                    "'{23}','{24}','{25}','{26}','{27}'," +
                    "'{28}','{29}','{30}','{31}'," +
                    "'{32}','{33}'," +
                    "'{34}','{35}','{36}','{37}','{38}'," +
                    "'{39}','{40}', '{41}','{42}', '{43}')",
                    dr["RegNum"].ToString(), dr["Name"].ToString(), dr["SecondName"].ToString(), dr["Surname"].ToString(),
                    dr["Privileges"].ToString(), QueryServ.QueryForBool(dr["IsExcellent"].ToString()), dr["ListenerTypeId"].ToString(), QueryServ.QueryForBool(dr["IsListener"].ToString()),
                    QueryServ.QueryForBool(dr["HostelEduc"].ToString()), dr["FacultyId"].ToString(), profId, specId,
                    dr["StudyBasisId"].ToString(), dr["StudyFormId"].ToString(), dr["CompetitionId"].ToString(),
                    dr["DocDate"].ToString(), _dRegion[dr["Nation"].ToString()], dr["RegionId"].ToString().Length == 0 ? "0" : dr["RegionId"].ToString(), dr["LanguageId"].ToString().Length == 0 ? "0" : dr["LanguageId"].ToString(),
                    educSeries, dr["AttestatRegion"].ToString(), educNum, QueryServ.QueryForBool(dr["CopyAtt"].ToString()),
                    dr["SchoolName"].ToString().Replace("'", ""), dr["SchoolCity"].ToString(), dr["SchoolNum"].ToString(), dr["SchoolTypeId"].ToString(), educYear,
                    ph, zc, a, la,
                    dr["BirthDate"].ToString(), QueryServ.QueryForBool(dr["Sex"].ToString()),
                    dr["PassportTypeId"].ToString(), dr["PassportSeries"].ToString(), dr["PassportNumber"].ToString(), dr["PassportDate"].ToString(), pa,
                    dr["StudyNumber"].ToString(), _slIds[dr["Id"].ToString()], dr["ObrazProgramName"].ToString(), dr["ObrazProgramCrypt"].ToString(),dr["StudyPlanNumber"].ToString());
                     
                _alQueries.Add(s);

                _NewId++;
            }
        }        

        //фильтр по факультету
        private string GetFilter(string table)
        {
            string res = string.Empty;

            if (!string.IsNullOrEmpty(FacultyId))
                res += string.Format(" AND {0}.FacultyId={1} ", table, FacultyId);
            
            res += string.Format(" AND {0}.StudyLevelGroupId={1} ", table, MainClass.studyLevelGroupId);

            return res;
        }

        private void btnMetro_Click(object sender, EventArgs e)
        {
            string newfile = folderBrowser.SelectedPath + "/MetroDB.mdb";

            FileInfo fi = new FileInfo(_metroBase);
            fi.CopyTo(newfile, true);

            _alQueries = new ArrayList();

            _odc = new OleDbClass();
            _odc.OpenDatabase(newfile);

            string query = string.Format("SELECT DISTINCT ed.extAbit.Id, ed.Person.Name, ed.Person.SecondName, ed.Person.Surname, " +
                              "ed.Person.BirthDate, ed.extAbit.StudyNumber, ed.extAbit.StudyLevelId, " +
                              "ed.Person.PassportTypeId, case when ed.Person.PassportTypeId=1 then 'Р' when ed.Person.PassportTypeId=3 then 'З' else '' end as PassportType, " +
                              "ed.Person.PassportSeries, ed.Person.PassportNumber,  " +
                              "ed.extEntryView.Id AS EntryProtId " +
                              "FROM ed.extAbit INNER JOIN ed.Person ON ed.extAbit.PersonId = ed.Person.Id " +
                              "INNER JOIN ed.extEntryView ON ed.extEntryView.AbiturientId = ed.extAbit.Id " +
                              "WHERE ed.extAbit.StudyFormId = 1 {0}", GetFilter("extAbit"));

            DataSet ds = _bdc.GetDataSet(query);

            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                string ser = dr["PassportSeries"].ToString();
                string ser1 = string.Empty, ser2 = string.Empty;
                string pType = dr["PassportTypeId"].ToString();

                int num = 0;
                string temp = ser.Replace(" ", "").Replace("-", "");
                if (pType == "1" && int.TryParse(temp, out num) && temp.Length == 4)
                {
                    ser = ser.Replace(" ", "").Replace("-", "");
                    ser1 = ser.Substring(0, 2);
                    ser2 = ser.Substring(2, 2);
                }
                else
                    ser1 = ser;

                string dateEnd;
                string course;

                string stLevel = dr["StudyLevelId"].ToString();
                if (stLevel == "16")
                {
                    dateEnd = "31.08.2016";
                    course = "1";
                }
                else if (stLevel == "17")
                {
                    dateEnd = "31.08.2014";
                    course = "5";
                }
                else
                {
                    dateEnd = "31.08.2017";
                    course = "1";
                }

                string datebirth = ((DateTime)dr["BirthDate"]).ToString("dd.MM.yyyy");

                string s = string.Format(
                    "INSERT INTO sList ([DOC_KIND], [DOC_SN], [DOC_S]," +
                    "[DOC_NUM],[SDOCUM],[NAME_F],[NAME_I], [NAME_O]," +
                    "[ORGCODE], [DATEEND], [BIRTHDAY], [COURSE])" +
                    "VALUES ('{0}','{1}','{2}'," +
                    "'{3}','{4}','{5}','{6}','{7}'," +
                    "'21','{8}','{9}','{10}')",
                    dr["PassportType"].ToString(), ser1, ser2,
                    dr["PassportNumber"].ToString(), dr["StudyNumber"].ToString(), dr["Surname"].ToString(), dr["Name"].ToString(), dr["SecondName"].ToString(),
                    dateEnd, datebirth, course);

                _alQueries.Add(s);
            }

            _odc.ExecuteWithTrasaction(_alQueries);
            MessageBox.Show("Done!");
        }
    }
}