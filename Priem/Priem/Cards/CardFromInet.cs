using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Data.Objects;
using System.Transactions;

using BaseFormsLib;
using EducServLib;

namespace Priem
{
    public partial class CardFromInet : CardFromList
    {        
        private DBPriem _bdcInet;    
        private int _abitBarc;
        private int _personBarc;

        private Guid? personId;
        private bool _closePerson;
        private bool _closeAbit;

        private List<ShortCompetition> LstCompetitions;

        LoadFromInet load;

        private DocsClass _docs;   

        // конструктор формы
        public CardFromInet(int abitBarcode, bool closeAbit, BaseFormEx fOwner, int? fOwnerRwInd)
        {
            InitializeComponent();
            _Id = null;
           
            _abitBarc = abitBarcode;
            _closeAbit = closeAbit;
            tcCard = tabCard;

            if (fOwner != null)
                this.formOwner = fOwner;

            if (fOwnerRwInd.HasValue)
                this.ownerRowIndex = fOwnerRwInd.Value;

            InitControls();     
        }      

        protected override void ExtraInit()
        { 
            base.ExtraInit();

            load = new LoadFromInet();
            _bdcInet = load.BDCInet;
            
            _bdc = MainClass.Bdc;
            _isModified = true;

            _personBarc = (int)_bdcInet.GetValue(string.Format("SELECT Person.Barcode FROM qAbiturient INNER JOIN Person ON qAbiturient.PersonId = Person.Id WHERE qAbiturient.CommitNumber = {0}", _abitBarc));

            lblBarcode.Text = _personBarc.ToString();
            lblBarcode.Text += @"\" + _abitBarc.ToString();

            _docs = new DocsClass(_personBarc, _abitBarc);

            tbNum.Enabled = false;

            rbMale.Checked = true;
            chbEkvivEduc.Visible = false;

            chbHostelAbitYes.Checked = false;
            chbHostelAbitNo.Checked = false;
            cbHEQualification.DropDownStyle = ComboBoxStyle.DropDown;
            
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbPassportType, HelpClass.GetComboListByTable("ed.PassportType"), true, false);
                    ComboServ.FillCombo(cbCountry, HelpClass.GetComboListByTable("ed.ForeignCountry", "ORDER BY LevelOfUsing DESC, Name"), true, false);
                    ComboServ.FillCombo(cbNationality, HelpClass.GetComboListByTable("ed.ForeignCountry", "ORDER BY LevelOfUsing DESC, Name"), true, false);
                    ComboServ.FillCombo(cbRegion, HelpClass.GetComboListByTable("ed.Region", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbLanguage, HelpClass.GetComboListByTable("ed.Language"), true, false);
                    ComboServ.FillCombo(cbCountryEduc, HelpClass.GetComboListByTable("ed.ForeignCountry", "ORDER BY LevelOfUsing DESC, Name"), true, false);                    
                    ComboServ.FillCombo(cbHEStudyForm, HelpClass.GetComboListByTable("ed.StudyForm"), true, false);

                    cbSchoolCity.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.SchoolCity AS Name FROM ed.Person_EducationInfo WHERE ed.Person_EducationInfo.SchoolCity > '' ORDER BY 1");
                    cbAttestatSeries.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.AttestatSeries AS Name FROM ed.Person_EducationInfo WHERE ed.Person_EducationInfo.AttestatSeries > '' ORDER BY 1");
                    cbHEQualification.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.HEQualification AS Name FROM ed.Person_EducationInfo WHERE NOT ed.Person_EducationInfo.HEQualification IS NULL AND ed.Person_EducationInfo.HEQualification > '' ORDER BY 1");

                    cbAttestatSeries.SelectedIndex = -1;
                    cbSchoolCity.SelectedIndex = -1;
                    cbHEQualification.SelectedIndex = -1;
                    
                    ComboServ.FillCombo(cbLanguage, HelpClass.GetComboListByTable("ed.Language"), true, false);

                    chbHostelEducYes.Checked = false;
                    chbHostelEducNo.Checked = false;
                }               

                // магистратура!
                if (MainClass.dbType == PriemType.PriemMag)
                {
                    ComboServ.FillCombo(cbSchoolType, HelpClass.GetComboListByQuery("SELECT Cast(ed.SchoolType.Id as nvarchar(100)) AS Id, ed.SchoolType.Name FROM ed.SchoolType WHERE ed.SchoolType.Id = 4 ORDER BY 1"), true, false);
                    tbSchoolNum.Visible = false;
                    tbSchoolName.Width = 200;
                    lblSchoolNum.Visible = false;
                    gbAtt.Visible = false;
                    gbDipl.Visible = true;
                    chbIsExcellent.Text = "Диплом с отличием";
                    btnAttMarks.Visible = false;
                    gbSchool.Visible = false;                   

                    gbEduc.Location = new Point(11, 7);
                    gbFinishStudy.Location = new Point(11, 222);

                }
                else
                {
                    tpDocs.Parent = null;
                    ComboServ.FillCombo(cbSchoolType, HelpClass.GetComboListByTable("ed.SchoolType", "ORDER BY 1"), true, false);                        
                }

                if (_closeAbit)
                    tpApplication.Parent = null;
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы " + exc.Message);
            }
        }

        protected override bool IsForReadOnly()
        {
            return !MainClass.RightsToEditCards();
        }
        
        #region handlers

        //инициализация обработчиков мегакомбов
        protected override void InitHandlers()
        {
            cbSchoolType.SelectedIndexChanged += new EventHandler(UpdateAfterSchool);
            cbCountry.SelectedIndexChanged += new EventHandler(UpdateAfterCountry);
            cbCountryEduc.SelectedIndexChanged += new EventHandler(UpdateAfterCountryEduc);
        }

        protected override void NullHandlers()
        {
            cbSchoolType.SelectedIndexChanged -= new EventHandler(UpdateAfterSchool);
            cbCountry.SelectedIndexChanged -= new EventHandler(UpdateAfterCountry);
            cbCountryEduc.SelectedIndexChanged -= new EventHandler(UpdateAfterCountryEduc);
        }

        private void UpdateAfterSchool(object sender, EventArgs e)
        {
            if (SchoolTypeId == MainClass.educSchoolId)
            {
                gbAtt.Visible = true;
                gbDipl.Visible = false;
                tbSchoolName.Width = 217;
            }
            else
            {
                if (SchoolTypeId == 4)
                    tbSchoolName.Width = 281;
                else
                    tbSchoolName.Width = 217;
                gbAtt.Visible = false;
                gbDipl.Visible = true;
            }
        }
        private void UpdateAfterCountry(object sender, EventArgs e)
        {
            if (CountryId == MainClass.countryRussiaId_Foreign)
            {
                cbRegion.Enabled = true;
                cbRegion.SelectedItem = "нет";
            }
            else
            {
                cbRegion.Enabled = false;
                cbRegion.SelectedItem = "нет";
            }
        }
        private void UpdateAfterCountryEduc(object sender, EventArgs e)
        {
            if (CountryEducId == MainClass.countryRussiaId_Foreign)
                chbEkvivEduc.Visible = false;
            else
                chbEkvivEduc.Visible = true;
        }
        private void chbHostelAbitYes_CheckedChanged(object sender, EventArgs e)
        {
            chbHostelAbitNo.Checked = !chbHostelAbitYes.Checked;
        }
        private void chbHostelAbitNo_CheckedChanged(object sender, EventArgs e)
        {
            chbHostelAbitYes.Checked = !chbHostelAbitNo.Checked;
        }

        #endregion

        protected override void FillCard()
        {
            try
            {
                FillPersonData(GetPerson());
                FillApplication();
                FillFiles();
            }
            catch (DataException de)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + de.Message);
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + ex.Message);
            }
        }

        private void FillFiles()
        {
            List<KeyValuePair<string, string>> lstFiles = _docs.UpdateFiles();
            if (lstFiles == null || lstFiles.Count == 0)
                return;

            chlbFile.DataSource = new BindingSource(lstFiles, null);
            chlbFile.ValueMember = "Key";
            chlbFile.DisplayMember = "Value";   
        }

        private extPerson GetPerson()
        {
            if (_personBarc == null)
                return null;

            try
            {
                if (!MainClass.CheckPersonBarcode(_personBarc))
                {
                    _closePerson = true;

                    using (PriemEntities context = new PriemEntities())
                    {
                        extPerson person = (from pers in context.extForeignPerson
                                            where pers.Barcode == _personBarc
                                            select pers).FirstOrDefault();

                        personId = person.Id;

                        tbNum.Text = person.PersonNum.ToString();
                        this.Text = "ПРОВЕРКА ДАННЫХ " + person.FIO;
                        
                        return person;
                    }
                }
                else
                {
                    if (_personBarc == 0)
                        return null;

                    _closePerson = false;
                    personId = null;

                    tcCard.SelectedIndex = 0;
                    tbSurname.Focus();

                    extPerson person = load.GetPersonByBarcode(_personBarc); 
                    
                    this.Text = "ЗАГРУЗКА " + person.FIO;
                    return person;
                }
            }

            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + ex.Message);
                return null;
            }       
        }

        private void FillPersonData(extPerson person)
        {
            if (person == null)
            {
                WinFormsServ.Error("Не найдены записи!");
                _isModified = false;
                this.Close();
            }

            try
            {
                PersonName = person.Name;
                SecondName = person.SecondName;
                Surname = person.Surname;
                BirthDate = person.BirthDate;
                BirthPlace = person.BirthPlace;
                PassportTypeId = person.PassportTypeId;
                PassportSeries = person.PassportSeries;
                PassportNumber = person.PassportNumber;
                PassportAuthor = person.PassportAuthor;
                PassportDate = person.PassportDate;
                PassportCode = person.PassportCode;
                PersonalCode = person.PersonalCode;
                Sex = person.Sex;
                CountryId = person.CountryId;
                NationalityId = person.NationalityId;
                RegionId = person.RegionId;
                Phone = person.Phone;
                Mobiles = person.Mobiles;
                Email = person.Email;
                Code = person.Code;
                City = person.City;
                Street = person.Street;
                House = person.House;
                Korpus = person.Korpus;
                Flat = person.Flat;
                HostelAbit = person.HostelAbit ?? false;                
                IsExcellent = person.IsExcellent ?? false;
                LanguageId = person.LanguageId;
                SchoolCity = person.SchoolCity;
                SchoolTypeId = person.SchoolTypeId;
                SchoolName = person.SchoolName;
                SchoolNum = person.SchoolNum;
                SchoolExitYear = person.SchoolExitYear;
                CountryEducId = person.CountryEducId;
                HasEkvivEduc = person.HasEkvivEduc ?? false;
                AttestatRegion = person.AttestatRegion;
                AttestatSeries = person.AttestatSeries;
                AttestatNum = person.AttestatNum;
                DiplomSeries = person.DiplomSeries;
                DiplomNum = person.DiplomNum;
                SchoolAVG = person.SchoolAVG;
                HighEducation = person.HighEducation;
                HEProfession = person.HEProfession;
                HEQualification = person.HEQualification;
                HEEntryYear = person.HEEntryYear;
                HEExitYear = person.HEExitYear;
                HEWork = person.HEWork;
                HEStudyFormId = person.HEStudyFormId;
                Stag = person.Stag;
                WorkPlace = person.WorkPlace;
                Privileges = person.Privileges;
                ExtraInfo = person.ExtraInfo;
                PersonInfo = person.PersonInfo;
                ScienceWork = person.ScienceWork;
                StartEnglish = person.StartEnglish ?? false;
                EnglishMark = person.EnglishMark;
            }
            catch (DataException de)
            {
                WinFormsServ.Error("Ошибка при заполнении формы (DataException)" + de.Message);
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + ex.Message);
            } 
        }

        public void FillApplication()
        {
            try
            {
                string query =
@"SELECT Abiturient.[Id]
,[PersonId]
,[Priority]
,[Barcode]
,[DateOfStart]
,[EntryId]
,[FacultyId]
,[FacultyName]
,[LicenseProgramId]
,[LicenseProgramCode]
,[LicenseProgramName]
,[ObrazProgramId]
,[ObrazProgramCrypt]
,[ObrazProgramName]
,[ProfileId]
,[ProfileName]
,[StudyBasisId]
,[StudyBasisName]
,[StudyFormId]
,[StudyFormName]
,[StudyLevelId]
,[StudyLevelName]
,[IsSecond]
,[IsReduced]
,[IsParallel]
,[CommitId]
,(SELECT MAX(ApplicationCommitVersion.Id) FROM ApplicationCommitVersion WHERE ApplicationCommitVersion.CommitId = [Abiturient].CommitId) AS VersionNum
,ApplicationCommit.IntNumber
FROM [Abiturient] 
INNER JOIN ApplicationCommit ON ApplicationCommit.Id = Abiturient.CommitId
WHERE IsCommited = 1 AND IntNumber=@CommitId";
                
                DataTable tbl = _bdcInet.GetDataSet(query, new SortedList<string, object>() { { "@CommitId", _abitBarc } }).Tables[0];

                LstCompetitions = 
                         (from DataRow rw in tbl.Rows
                          select new ShortCompetition(rw.Field<Guid>("Id"), rw.Field<Guid>("CommitId"), rw.Field<Guid>("EntryId"), rw.Field<Guid>("PersonId"))
                          {
                              CompetitionId = rw.Field<int>("StudyBasisId") == 1 ? 4 : 3,
                              CompetitionName = "не указана",
                              HasCompetition = false,
                              LicenseProgramId = rw.Field<int>("LicenseProgramId"),
                              LicenseProgramName = rw.Field<string>("LicenseProgramName"),
                              ObrazProgramId = rw.Field<int>("ObrazProgramId"),
                              ObrazProgramName = rw.Field<string>("ObrazProgramName"),
                              ProfileId = rw.Field<Guid?>("ProfileId"),
                              ProfileName = rw.Field<string>("ProfileName"),
                              StudyBasisId = rw.Field<int>("StudyBasisId"),
                              StudyBasisName = rw.Field<string>("StudyBasisName"),
                              StudyFormId = rw.Field<int>("StudyFormId"),
                              StudyFormName = rw.Field<string>("StudyFormName"),
                              StudyLevelId = rw.Field<int>("StudyLevelId"),
                              StudyLevelName = rw.Field<string>("StudyLevelName"),
                              FacultyId = rw.Field<int>("FacultyId"),
                              FacultyName = rw.Field<string>("FacultyName"),
                          }).ToList();

                UpdateApplicationGrid();

                //if (_closeAbit || _abitBarc == null)
                //    return;
                               
                //qAbiturient abit = load.GetAbitByBarcode(_abitBarc.Value);
                
                //if (abit == null)
                //{
                //    WinFormsServ.Error("Заявления отсутствуют!");
                //    _isModified = false;
                //    this.Close();
                //}

                //IsSecond = abit.IsSecond;
                //LicenseProgramId = abit.LicenseProgramId;
                //ObrazProgramId = abit.ObrazProgramId;
                //ProfileId = abit.ProfileId;
                //FacultyId = abit.FacultyId;
                //StudyFormId = abit.StudyFormId;
                //StudyBasisId = abit.StudyBasisId;
                //HostelEduc = abit.HostelEduc;
                //DocDate = abit.DocDate;
                //Priority = abit.Priority;
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы заявления" + ex.Message);
            }
        }

        private void UpdateApplicationGrid()
        {
            dgvApplications.DataSource = LstCompetitions.Select(x => new { x.Id, x.LicenseProgramName, x.ObrazProgramName, x.ProfileName, x.StudyFormName, x.StudyBasisName }).ToList();
            dgvApplications.Columns["Id"].Visible = false;
            dgvApplications.Columns["LicenseProgramName"].HeaderText = "Направление";
            dgvApplications.Columns["ObrazProgramName"].HeaderText = "Образ. программа";
            dgvApplications.Columns["ProfileName"].HeaderText = "Профиль";
            dgvApplications.Columns["StudyFormName"].HeaderText = "Форма обуч";
            dgvApplications.Columns["StudyBasisName"].HeaderText = "Основа обуч";
        }

        protected override void SetReadOnlyFieldsAfterFill()
        {
            base.SetReadOnlyFieldsAfterFill();
            
            if (_closePerson)
            {
                tcCard.SelectedTab = tpApplication;

                foreach (TabPage tp in tcCard.TabPages)
                {
                    if (tp != tpApplication && tp != tpDocs)
                    {
                        foreach (Control control in tp.Controls)
                        {
                            control.Enabled = false;
                            foreach (Control crl in control.Controls)
                                crl.Enabled = false;
                        }
                    }
                }
            }
        }

        #region Save

        //Вроде как эта функция пережёвывает апострофы в понятные серверу символы. Типа как mysql_escape_string()
        private string GetOConnorString(string source)
        {
            string res = "";
            foreach (char ch in source.ToCharArray())
            {
                if (ch == '\'')
                    res += '\\' + '\'';
                else
                    res += ch;
            }
            return res;
        }

        // проверка на уникальность абитуриента
        private bool CheckIdent()
        {
            using (PriemEntities context = new PriemEntities())
            {
                ObjectParameter boolPar = new ObjectParameter("result", typeof(bool));

                if(_Id == null)
                    context.CheckPersonIdent(GetOConnorString(Surname), PersonName, SecondName, BirthDate, PassportSeries, PassportNumber, AttestatRegion, AttestatSeries, AttestatNum, boolPar);
                else
                    context.CheckPersonIdentWithId(GetOConnorString(Surname), PersonName, SecondName, BirthDate, PassportSeries, PassportNumber, AttestatRegion, AttestatSeries, AttestatNum, GuidId, boolPar);

                return Convert.ToBoolean(boolPar.Value);
            }
        }

        protected override bool CheckFields()
        {
            if (Surname.Length <= 0)
            {
                epError.SetError(tbSurname, "Отсутствует фамилия абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (PersonName.Length <= 0)
            {
                epError.SetError(tbName, "Отсутствует имя абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            //Для О'Коннор сделал добавку в регулярное выражение: \'
            if (!Regex.IsMatch(Surname, @"^[А-Яа-яёЁ\-\'\s]+$"))
            {
                epError.SetError(tbSurname, "Неправильный формат");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!Regex.IsMatch(PersonName, @"^[А-Яа-яёЁ\-\s]+$"))
            {
                epError.SetError(tbName, "Неправильный формат");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!Regex.IsMatch(SecondName, @"^[А-Яа-яёЁ\-\s]*$"))
            {
                epError.SetError(tbSecondName, "Неправильный формат");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (SecondName.StartsWith("-"))
            {
                SecondName = SecondName.Replace("-", "");
            }

            // проверка на англ. буквы
            if (!Util.IsRussianString(PersonName))
            {
                epError.SetError(tbName, "Имя содержит английские символы, используйте только русскую раскладку");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!Util.IsRussianString(Surname))
            {
                epError.SetError(tbSurname, "Фамилия содержит английские символы, используйте только русскую раскладку");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!Util.IsRussianString(SecondName))
            {
                epError.SetError(tbSecondName, "Отчество содержит английские символы, используйте только русскую раскладку");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (BirthDate == null)
            {
                epError.SetError(dtBirthDate, "Неправильно указана дата");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            int checkYear = DateTime.Now.Year - 12;
            if (BirthDate.Value.Year > checkYear || BirthDate.Value.Year < 1920)
            {
                epError.SetError(dtBirthDate, "Неправильно указана дата");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (PassportDate.Value.Year > DateTime.Now.Year || PassportDate.Value.Year < 1970)
            {
                epError.SetError(dtPassportDate, "Неправильно указана дата");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (PassportTypeId == MainClass.pasptypeRFId)
            {
                if (!(PassportSeries.Length == 4))
                {
                    epError.SetError(tbPassportSeries, "Неправильно введена серия паспорта РФ абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epError.Clear();

                if (!(PassportNumber.Length == 6))
                {
                    epError.SetError(tbPassportNumber, "Неправильно введен номер паспорта РФ абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epError.Clear();
            }

            if (NationalityId == MainClass.countryRussiaId)
            {
                if (PassportSeries.Length <= 0)
                {
                    epError.SetError(tbPassportSeries, "Отсутствует серия паспорта абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epError.Clear();

                if (PassportNumber.Length <= 0)
                {
                    epError.SetError(tbPassportNumber, "Отсутствует номер паспорта абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epError.Clear();
            }

            if (PassportSeries.Length > 10)
            {
                epError.SetError(tbPassportSeries, "Слишком длинное значение серии паспорта абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();


            if (PassportNumber.Length > 20)
            {
                epError.SetError(tbPassportNumber, "Слишком длинное значение номера паспорта абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epError.Clear();

            if (!chbHostelAbitYes.Checked && !chbHostelAbitNo.Checked)
            {
                epError.SetError(chbHostelAbitNo, "Не указаны данные о предоставлении общежития");
                tabCard.SelectedIndex = 1;
                return false;
            }
            else
                epError.Clear();

            if (!Regex.IsMatch(SchoolExitYear.ToString(), @"^\d{0,4}$"))
            {
                epError.SetError(tbSchoolExitYear, "Неправильно указан год");
                tabCard.SelectedIndex = 2;
                return false;
            }
            else
                epError.Clear();

            if (gbAtt.Visible && AttestatNum.Length <= 0)
            {
                epError.SetError(tbAttestatNum, "Отсутствует номер аттестата абитуриента");
                tabCard.SelectedIndex = 2;
                return false;
            }
            else
                epError.Clear();

            double d = 0;
            if (tbSchoolAVG.Text.Trim() != "")
            {
                if (!double.TryParse(tbSchoolAVG.Text.Trim().Replace(".", ","), out d))
                {
                    epError.SetError(tbSchoolAVG, "Неправильный формат");
                    tabCard.SelectedIndex = 2;
                    return false;
                }
                else
                    epError.Clear();
            }

            //if (tbHEProfession.Text.Length >= 100)
            //{
            //    epError.SetError(tbHEProfession, "Длина поля превышает 100 символов.");
            //    tabCard.SelectedIndex = 2;
            //    return false;
            //}
            //else
            //    epError.Clear();

            if (tbScienceWork.Text.Length >= 2000)
            {
                epError.SetError(tbScienceWork, "Длина поля превышает 2000 символов. Укажите только самое основное.");
                tabCard.SelectedIndex = MainClass.dbType == PriemType.Priem ? 4 : 3;
                return false;
            }
            else
                epError.Clear();

            if (tbExtraInfo.Text.Length >= 1000)
            {
                epError.SetError(tbExtraInfo, "Длина поля превышает 1000 символов. Укажите только самое основное.");
                tabCard.SelectedIndex = MainClass.dbType == PriemType.Priem ? 4 : 3;
                return false;
            }
            else
                epError.Clear();

            if (tbPersonInfo.Text.Length > 1000)
            {
                epError.SetError(tbPersonInfo, "Длина поля превышает 1000 символов. Укажите только самое основное.");
                tabCard.SelectedIndex = MainClass.dbType == PriemType.Priem ? 4 : 3;
                return false;
            }
            else
                epError.Clear();

            if (tbWorkPlace.Text.Length > 1000)
            {
                epError.SetError(tbWorkPlace, "Длина поля превышает 1000 символов. Укажите только самое основное.");
                tabCard.SelectedIndex = MainClass.dbType == PriemType.Priem ? 4 : 3;
                return false;
            }
            else
                epError.Clear();

            if (!CheckIdent())
            {
                WinFormsServ.Error("В базе уже существует абитуриент с такими же либо ФИО, либо данными паспорта, либо данными аттестата!");
                return false;
            }

            return true;
        }

        private bool CheckFieldsAbit()
        {
            using (PriemEntities context = new PriemEntities())
            {
                if (!chbHostelEducYes.Checked && !chbHostelEducNo.Checked)
                {
                    epError.SetError(chbHostelEducNo, "Не указаны данные о предоставлении общежития");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epError.Clear();
            }

            return true;
        } 
        
        protected override bool SaveClick()
        {
            try
            {
                if (_closePerson)
                {
                    if (!SaveApplication())
                        return false;
                }
                else
                {
                    if (!CheckFields())
                        return false;

                    using (PriemEntities context = new PriemEntities())
                    {
                        using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                        {
                            try
                            {
                                ObjectParameter entId = new ObjectParameter("id", typeof(Guid));
                                context.Person_Foreign_insert(_personBarc, PersonName, SecondName, Surname, BirthDate, BirthPlace, PassportTypeId, PassportSeries, PassportNumber,
                                    PassportAuthor, PassportDate, Sex, CountryId, NationalityId, RegionId, Phone, Mobiles, Email, Code, City, Street, House, Korpus, Flat, HostelAbit, false,
                                    null, false, null, IsExcellent, LanguageId, SchoolCity, SchoolTypeId, SchoolName, SchoolNum, SchoolExitYear,
                                    SchoolAVG, CountryEducId, HasEkvivEduc, AttestatRegion, AttestatSeries, AttestatNum, DiplomSeries, DiplomNum, HighEducation, HEProfession,
                                    HEQualification, HEEntryYear, HEExitYear, HEStudyFormId, HEWork, Stag, WorkPlace, Privileges, PassportCode,
                                    PersonalCode, PersonInfo, ExtraInfo, ScienceWork, StartEnglish, EnglishMark, entId);

                                personId = (Guid)entId.Value;

                                transaction.Complete();
                            }
                            catch (Exception exc)
                            {
                                WinFormsServ.Error("Ошибка при сохранении:\n" + exc.Message + (exc.InnerException != null ? "\n\nВнутреннее исключение:\n" + exc.InnerException.Message : ""));
                            }
                        }
                        if (!SaveApplication())
                        {
                            _closePerson = true;
                            return false;
                        }
                        
                        _bdcInet.ExecuteQuery("UPDATE Person SET IsImported = 1 WHERE Person.Barcode = " + _personBarc);                       
                    }
                }  
                             
                _isModified = false;

                OnSave();               

                this.Close();
                return true;
            }
            catch (Exception de)
            {
                WinFormsServ.Error("Ошибка обновления данных" + de.Message);
                return false;
            }
        }

        private bool SaveApplication()
        {
            if (_closeAbit)
                return true;

            if (personId == null)
                return false;

            if (!CheckFieldsAbit())
                return false;

            try
            {
                using (TransactionScope trans = new TransactionScope())
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        ObjectParameter entId = new ObjectParameter("id", typeof(Guid));

                        //context.Abiturient_Insert(personId, EntryId, CompetitionId, HostelEduc, IsListener, WithHE, false, false, null, DocDate, DateTime.Now,
                        //AttDocOrigin, EgeDocOrigin, false, false, null, OtherCompetitionId, CelCompetitionId, CelCompetitionText, LanguageId, false,
                        //Priority, _abitBarc, entId);
                    }

                    _bdcInet.ExecuteQuery("UPDATE Application SET IsImported = 1 WHERE Application.Barcode = " + _abitBarc);

                    trans.Complete();
                    return true;
                }
            }
            catch (Exception de)
            {
                WinFormsServ.Error("Ошибка обновления данных Abiturient" + de.Message);
                return false;
            }
        }
       
        public bool IsMatchEgeNumber(string number)
        {
            string num = number.Trim();
            if (Regex.IsMatch(num, @"^\d{2}-\d{9}-(10|11|12)$"))//не даёт перегрузить воякам свои древние ЕГЭ, добавлен 2010 год
                return true;
            else
                return false;
        }

        #endregion 
        
        protected override void OnClosed()
        {
            base.OnClosed();
            load.CloseDB();                
        }

        protected override void OnSave()
        {
            base.OnSave();
            using (PriemEntities context = new PriemEntities())
            {
                if (_abitBarc != null)
                {
                    Guid? abId = (from ab in context.qAbiturient
                                  where ab.Barcode == _abitBarc
                                  select ab.Id).FirstOrDefault();

                    MainClass.OpenCardAbit(abId.ToString(), null, null);

                }
                else
                {
                    Guid? perId = (from per in context.extForeignPerson
                                   where per.Barcode == _personBarc
                                   select per.Id).FirstOrDefault();

                    MainClass.OpenCardPerson(perId.ToString(), null, null);
                }
            }
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            List<KeyValuePair<string, string>> lstFiles = new List<KeyValuePair<string, string>>();
            foreach (KeyValuePair<string, string> file in chlbFile.CheckedItems)
            {
                lstFiles.Add(file);
            }

            _docs.OpenFile(lstFiles);
        }

        private void tabCard_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.D1)
                this.tcCard.SelectedIndex = 0;
            if (e.Control && e.KeyCode == Keys.D2)
                this.tcCard.SelectedIndex = 1;
            if (e.Control && e.KeyCode == Keys.D3)
                this.tcCard.SelectedIndex = 2;
            if (e.Control && e.KeyCode == Keys.D4)
                this.tcCard.SelectedIndex = 3;
            if (e.Control && e.KeyCode == Keys.D5)
                this.tcCard.SelectedIndex = 4;
            if (e.Control && e.KeyCode == Keys.D6)
                this.tcCard.SelectedIndex = 5;
            if (e.Control && e.KeyCode == Keys.D7)
                this.tcCard.SelectedIndex = 6;
            if (e.Control && e.KeyCode == Keys.D8)
                this.tcCard.SelectedIndex = 7;
            if (e.Control && e.KeyCode == Keys.S)
                SaveRecord();
        }

        private void dgvApplications_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int rwNum = e.RowIndex;
            if (rwNum >= 0)
            {
                var ent = GetCompFromGrid(rwNum);
                if (ent != null)
                {
                    var crd = new CardCompetitionInInet(ent);
                    crd.Show();
                }
            }
        }

        private ShortCompetition GetCompFromGrid(int rwNum)
        {
            if (rwNum < 0)
                return null;

            Guid Id = (Guid)dgvApplications["Id", rwNum].Value;
            return LstCompetitions.Where(x => x.Id == Id).FirstOrDefault();
        }
    }
}
