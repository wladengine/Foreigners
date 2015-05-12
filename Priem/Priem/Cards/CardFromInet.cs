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
using PriemLib;

namespace Priem
{
    public partial class CardFromInet : CardFromList
    {
        #region Fields
        private DBPriem _bdcInet;
        private int? _abitBarc;
        private int? _personBarc;
        private int _currentEducRow;

        private Guid? personId;
        private bool _closePerson;
        private bool _closeAbit;

        PriemLib.LoadFromInet load;
        private List<ShortCompetition> LstCompetitions;
        private List<Person_EducationInfo> lstEducationInfo;

        private DocsClass _docs;
        #endregion

        // конструктор формы
        public CardFromInet(int? personBarcode, int? abitBarcode, bool closeAbit)
        {
            InitializeComponent();
            _Id = null;
           
            _abitBarc = abitBarcode;
            _personBarc = personBarcode;
            _closeAbit = closeAbit;
            tcCard = tabCard;
            
            if (_abitBarc == null)
                _closeAbit = true;

            InitControls();     
        }      

        protected override void ExtraInit()
        { 
            base.ExtraInit();

            load = new LoadFromInet();
            _bdcInet = load.BDCInet;
            
            _bdc = MainClass.Bdc;
            _isModified = true;

            if (_personBarc == null)
                _personBarc = (int)_bdcInet.GetValue(string.Format("SELECT Person.Barcode FROM Abiturient INNER JOIN Person ON Abiturient.PersonId = Person.Id WHERE Abiturient.ApplicationCommitNumber = {0}", _abitBarc));

            lblBarcode.Text = _personBarc.ToString();
            if (_abitBarc != null)
                lblBarcode.Text += @"\" + _abitBarc.ToString();

            _docs = new DocsClass(_personBarc.Value, _abitBarc);

            tbNum.Enabled = false;

            rbMale.Checked = true;
            chbEkvivEduc.Visible = false;

            chbHostelAbitYes.Checked = false;
            chbHostelAbitNo.Checked = false;
            chbHostelEducYes.Checked = false;
            chbHostelEducNo.Checked = false;

            cbHEQualification.DropDownStyle = ComboBoxStyle.DropDown;
            
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbPassportType, HelpClass.GetComboListByTable("ed.PassportType"), true, false);
                    ComboServ.FillCombo(cbCountry, HelpClass.GetComboListByTable("ed.Country", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbNationality, HelpClass.GetComboListByTable("ed.Country", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbRegion, HelpClass.GetComboListByTable("ed.Region", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbRegionEduc, HelpClass.GetComboListByTable("ed.Region", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbLanguage, HelpClass.GetComboListByTable("ed.Language"), true, false);
                    ComboServ.FillCombo(cbCountryEduc, HelpClass.GetComboListByTable("ed.Country", "ORDER BY Distance, Name"), true, false);                    
                    ComboServ.FillCombo(cbHEStudyForm, HelpClass.GetComboListByTable("ed.StudyForm"), true, false);
                    ComboServ.FillCombo(cbMSStudyForm, HelpClass.GetComboListByTable("ed.StudyForm"), true, false);
                    ComboServ.FillCombo(cbSchoolType, HelpClass.GetComboListByTable("ed.SchoolType", "ORDER BY 1"), true, false);

                    cbSchoolCity.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.SchoolCity AS Name FROM ed.Person_EducationInfo WHERE ed.Person_EducationInfo.SchoolCity > '' ORDER BY 1");
                    cbAttestatSeries.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.AttestatSeries AS Name FROM ed.Person_EducationInfo WHERE ed.Person_EducationInfo.AttestatSeries > '' ORDER BY 1");
                    cbHEQualification.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.HEQualification AS Name FROM ed.Person_EducationInfo WHERE NOT ed.Person_EducationInfo.HEQualification IS NULL AND ed.Person_EducationInfo.HEQualification > '' ORDER BY 1");

                    cbAttestatSeries.SelectedIndex = -1;
                    cbSchoolCity.SelectedIndex = -1;
                    cbHEQualification.SelectedIndex = -1;
                    
                    ComboServ.FillCombo(cbLanguage, HelpClass.GetComboListByTable("ed.Language"), true, false);
                }

                // ЕГЭ только для 1 курса!
                if (MainClass.dbType != PriemType.Priem)
                {
                    tpEge.Parent = null;
                    tpSecond.Parent = null;
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

            if (MainClass.dbType == PriemType.PriemMag)
            {
                btnSaveChange.Text = "Одобрить";
                btnSaveChange.Enabled = (MainClass.bMagImportApplicationsEnabled || MainClass.IsTestDB);
            }
        }

        #region handlers

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
            if (SchoolTypeId == 1)
            {
                gbAtt.Visible = true;
                gbDipl.Visible = false;
            }
            else
            {
                gbDipl.Visible = true;
                gbAtt.Visible = false;
            }
        }
        private void UpdateAfterCountry(object sender, EventArgs e)
        {
            if (CountryId == MainClass.countryRussiaId)
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
            if (CountryEducId == MainClass.countryRussiaId)
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

        #endregion

        #region Fill Card

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
                        extPerson person = (from pers in context.extPerson
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

                    extPerson person = load.GetPersonByBarcode(_personBarc.Value);

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
                SNILS = person.SNILS;
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
                CodeReal = person.CodeReal;
                CityReal = person.CityReal;
                StreetReal = person.StreetReal;
                HouseReal = person.HouseReal;
                KorpusReal = person.KorpusReal;
                FlatReal = person.FlatReal;
                KladrCode = person.KladrCode;
                HostelAbit = person.HostelAbit;
                HostelEduc = person.HostelEduc;
                LanguageId = person.LanguageId;
                Stag = person.Stag;
                WorkPlace = person.WorkPlace;
                MSVuz = person.MSVuz;
                MSCourse = person.MSCourse;
                MSStudyFormId = person.MSStudyFormId;
                Privileges = person.Privileges;
                ExtraInfo = person.ExtraInfo;
                PersonInfo = person.PersonInfo;
                ScienceWork = person.ScienceWork;
                StartEnglish = person.StartEnglish;
                EnglishMark = person.EnglishMark;

                FillEducationData(load.GetPersonEducationDocumentsByBarcode(_personBarc.Value));

                if (MainClass.dbType == PriemType.Priem)
                {
                    DataTable dtEge = load.GetPersonEgeByBarcode(_personBarc.Value);
                    FillEgeFirst(dtEge);
                }
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
        private void FillEgeFirst(DataTable dtEge)
        {
            if (MainClass.dbType == PriemType.PriemMag)
                return;

            try
            {
                DataTable examTable = new DataTable();

                DataColumn clm;
                clm = new DataColumn();
                clm.ColumnName = "Предмет";
                clm.ReadOnly = true;
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "ExamId";
                clm.ReadOnly = true;
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "Баллы";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "Номер сертификата";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "Типографский номер";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "EgeCertificateId";
                examTable.Columns.Add(clm);


                string defQuery = "SELECT ed.EgeExamName.Name AS 'Предмет', ed.EgeExamName.Id AS ExamId FROM ed.EgeExamName";
                DataSet ds = _bdc.GetDataSet(defQuery);
                foreach (DataRow dsRow in ds.Tables[0].Rows)
                {
                    DataRow newRow;
                    newRow = examTable.NewRow();
                    newRow["Предмет"] = dsRow["Предмет"].ToString();
                    newRow["ExamId"] = dsRow["ExamId"].ToString();
                    examTable.Rows.Add(newRow);
                }

                foreach (DataRow dsRow in dtEge.Rows)
                {
                    for (int i = 0; i < examTable.Rows.Count; i++)
                    {
                        if (examTable.Rows[i]["ExamId"].ToString() == dsRow["ExamId"].ToString())
                        {
                            examTable.Rows[i]["Баллы"] = dsRow["Value"].ToString();
                            examTable.Rows[i]["Номер сертификата"] = dsRow["Number"].ToString();
                        }
                    }
                }

                DataView dv = new DataView(examTable);
                dv.AllowNew = false;

                dgvEGE.DataSource = dv;
                dgvEGE.Columns["ExamId"].Visible = false;
                dgvEGE.Columns["EgeCertificateId"].Visible = false;

                dgvEGE.Columns["Предмет"].Width = 162;
                dgvEGE.Columns["Баллы"].Width = 45;
                dgvEGE.Columns["Номер сертификата"].Width = 110;
                dgvEGE.ReadOnly = false;

                dgvEGE.Update();
            }
            catch (DataException de)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + de.Message);
            }
        }
        private void FillFiles()
        {
            List<KeyValuePair<string, string>> lstFiles = _docs.UpdateFiles();
            if (lstFiles == null || lstFiles.Count == 0)
                return;

            dgvFiles.DataSource = _docs.UpdateFilesTable();
            if (dgvFiles.Rows.Count > 0)
            {
                foreach (DataGridViewColumn clm in dgvFiles.Columns)
                    clm.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                
                if (!dgvFiles.Columns.Contains("Открыть"))
                {
                    DataGridViewCheckBoxCell cl = new DataGridViewCheckBoxCell();
                    cl.TrueValue = true;
                    cl.FalseValue = false;

                    DataGridViewCheckBoxColumn clm = new DataGridViewCheckBoxColumn();
                    clm.CellTemplate = cl;
                    clm.Name = "Открыть";
                    dgvFiles.Columns.Add(clm);
                    dgvFiles.Columns["Открыть"].DisplayIndex = 0;
                    dgvFiles.Columns["Открыть"].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader; 
                }
                if (dgvFiles.Columns.Contains("Id"))
                    dgvFiles.Columns["Id"].Visible = false;
                if (dgvFiles.Columns.Contains("FileExtention"))
                    dgvFiles.Columns["FileExtention"].Visible = false;
                dgvFiles.Columns["FileName"].HeaderText = "Файл";
                dgvFiles.Columns["Comment"].HeaderText = "Комментарий";
                dgvFiles.Columns["FileTypeName"].HeaderText = "Тип файла";
                dgvFiles.Columns["FileName"].ReadOnly = true;
                dgvFiles.Columns["Comment"].ReadOnly = true;
                dgvFiles.Columns["FileTypeName"].ReadOnly = true;
            }
        }
        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            List<KeyValuePair<string, string>> lstFiles = new List<KeyValuePair<string, string>>();
            lstFiles = new List<KeyValuePair<string, string>>();
            foreach (DataGridViewRow rw in dgvFiles.Rows)
            {
                DataGridViewCheckBoxCell cell = rw.Cells["Открыть"] as DataGridViewCheckBoxCell;
                if (cell.Value == cell.TrueValue)
                {
                    if (dgvFiles.Columns.Contains("FileName"))
                    {
                        string fileName = rw.Cells["FileName"].Value.ToString(); 
                        KeyValuePair<string, string> file = new KeyValuePair<string, string>(rw.Cells["Id"].Value.ToString(), fileName);
                        lstFiles.Add(file);
                    }
                }
            }
            _docs.OpenFile(lstFiles);
        }
        private void btnDocCardOpen_Click(object sender, EventArgs e)
        {
            if (_personBarc != null)
                new PriemLib.DocCard(_personBarc.Value, null, false).Show();
        }

        #region Applications
        public void FillApplication()
        {
            try
            {
                LstCompetitions = load.GetCompetitionList(_abitBarc.Value);
                if (LstCompetitions.Count == 0)
                {
                    WinFormsServ.Error("Заявления отсутствуют!");
                    _isModified = false;
                    this.Close();
                }

                tbApplicationVersion.Text = (LstCompetitions[0].VersionNum.HasValue ? "№ " + LstCompetitions[0].VersionNum.Value.ToString() : "n/a") +
                    (LstCompetitions[0].VersionDate.HasValue ? (" от " + LstCompetitions[0].VersionDate.Value.ToShortDateString() + " " + LstCompetitions[0].VersionDate.Value.ToShortTimeString()) : "n/a");

                UpdateApplicationGrid();
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы заявления" + ex.Message);
            }
        }
        private void UpdateApplicationGrid()
        {
            dgvApplications.DataSource = LstCompetitions.OrderBy(x => x.Priority)
                .Select(x => new
                {
                    x.Id,
                    x.Priority,
                    x.LicenseProgramName,
                    x.ObrazProgramName,
                    x.ProfileName,
                    x.StudyFormName,
                    x.StudyBasisName,
                    HasCompetition = x.HasCompetition || x.IsApprovedByComission,
                    comp = x.lstInnerEntryInEntry.Count > 0 ? "есть приоритеты" : ""
                }).ToList();

            dgvApplications.Columns["Id"].Visible = false;
            dgvApplications.Columns["Priority"].HeaderText = "Приор";
            dgvApplications.Columns["Priority"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader;
            dgvApplications.Columns["LicenseProgramName"].HeaderText = "Направление";
            dgvApplications.Columns["ObrazProgramName"].HeaderText = "Образ. программа";
            dgvApplications.Columns["ProfileName"].HeaderText = "Профиль";
            dgvApplications.Columns["StudyFormName"].HeaderText = "Форма обуч";
            dgvApplications.Columns["StudyBasisName"].HeaderText = "Основа обуч";
            dgvApplications.Columns["comp"].HeaderText = "";
            dgvApplications.Columns["HasCompetition"].Visible = false;
        }
        private void dgvApplications_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if ((bool)dgvApplications["HasCompetition", e.RowIndex].Value)
                {
                    e.CellStyle.BackColor = Color.Cyan;
                    e.CellStyle.SelectionBackColor = Color.Cyan;
                }
                else
                {
                    e.CellStyle.BackColor = Color.Coral;
                    e.CellStyle.SelectionBackColor = Color.Coral;
                }
            }
        }
        private void dgvApplications_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int rwNum = e.RowIndex;
            OpenCardCompetitionInInet(rwNum);
        }
        private void btnOpenCompetition_Click(object sender, EventArgs e)
        {
            if (dgvApplications.SelectedCells.Count == 0)
                return;

            int rwNum = dgvApplications.SelectedCells[0].RowIndex;
            OpenCardCompetitionInInet(rwNum);
        }
        private ShortCompetition GetCompFromGrid(int rwNum)
        {
            if (rwNum < 0)
                return null;

            Guid Id = (Guid)dgvApplications["Id", rwNum].Value;
            return LstCompetitions.Where(x => x.Id == Id).FirstOrDefault();
        }
        private void OpenCardCompetitionInInet(int rwNum)
        {
            if (rwNum >= 0)
            {
                var ent = GetCompFromGrid(rwNum);
                if (ent != null)
                {
                    var crd = new CardCompetitionInInet(ent);
                    crd.OnUpdate += UpdateCommitCompetition;
                    crd.Show();
                }
            }
        }
        private void UpdateCommitCompetition(ShortCompetition comp)
        {
            int ind = LstCompetitions.FindIndex(x => comp.Id == x.Id);
            if (ind > -1)
            {
                LstCompetitions[ind].HasCompetition = true;
                LstCompetitions[ind].IsApprovedByComission = true;
                LstCompetitions[ind].CompetitionId = comp.CompetitionId;
                LstCompetitions[ind].CompetitionName = comp.CompetitionName;

                LstCompetitions[ind].DocInsertDate = comp.DocInsertDate;
                LstCompetitions[ind].IsGosLine = comp.IsGosLine;
                LstCompetitions[ind].IsListener = comp.IsListener;
                LstCompetitions[ind].IsReduced = comp.IsReduced;

                LstCompetitions[ind].FacultyId = comp.FacultyId;
                LstCompetitions[ind].FacultyName = comp.FacultyName;
                LstCompetitions[ind].LicenseProgramId = comp.LicenseProgramId;
                LstCompetitions[ind].LicenseProgramName = comp.LicenseProgramName;
                LstCompetitions[ind].ObrazProgramId = comp.ObrazProgramId;
                LstCompetitions[ind].ObrazProgramName = comp.ObrazProgramName;
                LstCompetitions[ind].ProfileId = comp.ProfileId;
                LstCompetitions[ind].ProfileName = comp.ProfileName;

                LstCompetitions[ind].StudyFormId = comp.StudyFormId;
                LstCompetitions[ind].StudyFormName = comp.StudyFormName;
                LstCompetitions[ind].StudyBasisId = comp.StudyBasisId;
                LstCompetitions[ind].StudyBasisName = comp.StudyBasisName;
                LstCompetitions[ind].StudyLevelId = comp.StudyLevelId;
                LstCompetitions[ind].StudyLevelName = comp.StudyLevelName;

                LstCompetitions[ind].HasCompetition = comp.HasCompetition;
                LstCompetitions[ind].ChangeEntry();

                load.UpdateApplicationSetApprovedByComission(comp);

                UpdateApplicationGrid();
            }
        }
        #endregion

        #region EducationInfo
        private void FillEducationData(List<Person_EducationInfo> lstVals)
        {
            lstEducationInfo = lstVals;

            dgvEducationDocuments.DataSource = lstVals.Select(x => new
            {
                x.Id,
                School = x.SchoolName,
                Series = (x.SchoolTypeId == 1 ? x.AttestatSeries : x.DiplomSeries),
                Num = x.SchoolTypeId == 1 ? x.AttestatNum : x.DiplomNum,
            }).ToList();

            dgvEducationDocuments.Columns["Id"].Visible = false;
            dgvEducationDocuments.Columns["School"].HeaderText = "Уч. учреждение";
            dgvEducationDocuments.Columns["School"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvEducationDocuments.Columns["Series"].HeaderText = "Серия";
            dgvEducationDocuments.Columns["Series"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dgvEducationDocuments.Columns["Num"].HeaderText = "Номер";
            dgvEducationDocuments.Columns["Num"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            if (lstVals.Count > 0)
                ViewEducationInfo(lstVals.First().Id);

            _currentEducRow = 0;
        }
        private void dgvEducationDocuments_CurrentCellChanged(object sender, EventArgs e)
        {
            if (dgvEducationDocuments.CurrentRow != null)
                if (dgvEducationDocuments.CurrentRow.Index != _currentEducRow)
                {
                    _currentEducRow = dgvEducationDocuments.CurrentRow.Index;
                    ViewEducationInfo(lstEducationInfo[_currentEducRow].Id);
                }
        }
        private void ViewEducationInfo(int id)
        {
            int ind = lstEducationInfo.FindIndex(x => x.Id == id);

            CountryEducId = lstEducationInfo[ind].CountryEducId;
            RegionEducId = lstEducationInfo[ind].RegionEducId;

            tbEqualityDocumentNumber.Visible = CountryEducId != MainClass.countryRussiaId;
            chbEkvivEduc.Visible = CountryEducId != MainClass.countryRussiaId;
            IsEqual = lstEducationInfo[ind].IsEqual;
            EqualDocumentNumber = lstEducationInfo[ind].EqualDocumentNumber;

            SchoolTypeId = lstEducationInfo[ind].SchoolTypeId;
            if (SchoolTypeId == 1)
            {
                gbAtt.Visible = true;
                gbDipl.Visible = false;
                AttestatSeries = lstEducationInfo[ind].AttestatSeries;
                AttestatNum = lstEducationInfo[ind].AttestatNum;

                gbFinishStudy.Visible = false;
                lblSchoolNum.Visible = true;
                tbSchoolNum.Visible = true;
                btnAttMarks.Visible = true;
                chbIsExcellent.Text = "Медалист (отличник)";
            }
            else
            {
                gbAtt.Visible = false;
                gbDipl.Visible = true;
                DiplomSeries = lstEducationInfo[ind].DiplomSeries;
                DiplomNum = lstEducationInfo[ind].DiplomNum;

                gbFinishStudy.Visible = true;
                lblSchoolNum.Visible = false;
                tbSchoolNum.Visible = false;
                btnAttMarks.Visible = false;
                chbIsExcellent.Text = "Диплом с отличием";

                HighEducation = lstEducationInfo[ind].HighEducation;
                HEProfession = lstEducationInfo[ind].HEProfession;
                HEQualification = lstEducationInfo[ind].HEQualification;
                HEEntryYear = lstEducationInfo[ind].HEEntryYear;
                HEExitYear = lstEducationInfo[ind].HEExitYear;
                HEWork = lstEducationInfo[ind].HEWork;
                HEStudyFormId = lstEducationInfo[ind].HEStudyFormId;
            }
            
            SchoolAVG = lstEducationInfo[ind].SchoolAVG;
            IsExcellent = lstEducationInfo[ind].IsExcellent;
            SchoolCity = lstEducationInfo[ind].SchoolCity;
            SchoolName = lstEducationInfo[ind].SchoolName;
            SchoolNum = lstEducationInfo[ind].SchoolNum;
            SchoolExitYear = lstEducationInfo[ind].SchoolExitYear;
        }
        #endregion

        #endregion

        #region Save

        // проверка на уникальность абитуриента
        private bool CheckIdent()
        {
            using (PriemEntities context = new PriemEntities())
            {
                ObjectParameter boolPar = new ObjectParameter("result", typeof(bool));

                if(_Id == null)
                    context.CheckPersonIdent(Surname, PersonName, SecondName, BirthDate, PassportSeries, PassportNumber, AttestatSeries, AttestatNum, boolPar);
                else
                    context.CheckPersonIdentWithId(Surname, PersonName, SecondName, BirthDate, PassportSeries, PassportNumber, AttestatSeries, AttestatNum, GuidId, boolPar);

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

            if (tbHEProfession.Text.Length >= 1000)
            {
                epError.SetError(tbHEProfession, "Длина поля превышает 1000 символов.");
                tabCard.SelectedIndex = 2;
                return false;
            }
            else
                epError.Clear();

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

            if (MainClass.dbType == PriemType.Priem)
            {
                SortedList<string, string> slNumbers = new SortedList<string, string>();

                foreach (DataGridViewRow dr in dgvEGE.Rows)
                {
                    string num = dr.Cells["Номер сертификата"].Value.ToString();
                    string prNum = dr.Cells["Типографский номер"].Value.ToString();
                    string balls = dr.Cells["Баллы"].Value.ToString();

                    if (num.Length == 0 && balls.Length == 0)
                        continue;

                    int bls;
                    if (!(int.TryParse(balls, out bls) && bls > 0 && bls < 101))
                    {
                        epError.SetError(dgvEGE, "Неверно введены баллы");
                        tabCard.SelectedIndex = 3;
                        return false;
                    }
                    else
                        epError.Clear();

                    if (!EgeDataProvider.GetIsMatchEgeNumber(num))
                    {
                        epError.SetError(dgvEGE, "Номер свидетельства не соответствует формату **-*********-**");
                        tabCard.SelectedIndex = 3;
                        return false;
                    }
                    else
                        epError.Clear();

                    if (slNumbers.Keys.Contains(num))
                    {
                        if (slNumbers[num].CompareTo(prNum) != 0)
                        {
                            epError.SetError(dgvEGE, "У свидетельств с одним номером разные типографские номера");
                            tabCard.SelectedIndex = 3;
                            return false;
                        }
                    }
                    else
                    {
                        epError.Clear();
                        slNumbers.Add(num, prNum);
                    }
                }
            }

            return true;
        }
        
        protected override bool SaveClick()
        {
            try
            {
                if (_closePerson)
                {
                    if (!SaveApplication(personId.Value))
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
                                context.Person_insert(_personBarc, PersonName, SecondName, Surname, BirthDate, BirthPlace, PassportTypeId, PassportSeries, PassportNumber,
                                    PassportAuthor, PassportDate, Sex, CountryId, NationalityId, RegionId, Phone, Mobiles, Email,
                                    Code, City, Street, House, Korpus, Flat, CodeReal, CityReal, StreetReal, HouseReal, KorpusReal, FlatReal, KladrCode, HostelAbit, HostelEduc, false,
                                    null, false, null, LanguageId, Stag, WorkPlace, MSVuz, MSCourse, MSStudyFormId, Privileges, PassportCode,
                                    PersonalCode, PersonInfo, ExtraInfo, ScienceWork, StartEnglish, EnglishMark, EgeInSpbgu, SNILS, HasTRKI, TRKICertificateNumber, entId);

                                personId = (Guid)entId.Value;

                                SaveEducationDocuments();
                                SaveEgeFirst();
                                transaction.Complete();
                            }
                            catch (Exception exc)
                            {
                                WinFormsServ.Error("Ошибка при сохранении:", exc);
                            }
                        }
                        if (!SaveApplication(personId.Value))
                        {
                            _closePerson = true;
                            return false;
                        }
                        
                        if (!MainClass.IsTestDB)
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

        private bool SaveApplication(Guid PersonId)
        {
            if (_closeAbit)
                return true;

            if (personId == null)
                return false;

            //if (!CheckFieldsAbit())
            //    return false;

            try
            {
                using (TransactionScope trans = new TransactionScope(TransactionScopeOption.Required))
                {
                    ApplicationCommitSaveProvider.CheckAndUpdateNotUsedApplications(personId.Value, LstCompetitions);
                    ApplicationCommitSaveProvider.SaveApplicationCommitInWorkBase(personId.Value, LstCompetitions, LanguageId, _abitBarc);
                    
                    trans.Complete();
                }
                
                load.UpdateApplicationCommitSetIsImported(_abitBarc);
                return true;
            }
            catch (Exception de)
            {
                WinFormsServ.Error("Ошибка обновления данных Abiturient\n" + de.Message + "\n" + de.InnerException.Message);
                return false;
            }
        }
        private void SaveEgeFirst()
        {
            if (MainClass.dbType == PriemType.PriemMag)
                return;

            try
            {
                EgeList egeLst = new EgeList();
                foreach (DataGridViewRow dr in dgvEGE.Rows)
                {
                    if (dr.Cells["Баллы"].Value.ToString().Trim() != string.Empty)
                        egeLst.Add(new EgeMarkCert(dr.Cells["ExamId"].Value.ToString().Trim(), dr.Cells["Баллы"].Value.ToString().Trim(), dr.Cells["Номер сертификата"].Value.ToString().Trim(), dr.Cells["Типографский номер"].Value.ToString()));
                }

                EgeDataProvider.SaveEgeFromEgeList(egeLst, personId.Value);
            }
            catch (Exception de)
            {          
                WinFormsServ.Error("Ошибка сохранения данные ЕГЭ - данные не были сохранены. Введите их заново! \n" + de.Message);
            }
        }
        private void SaveEducationDocuments()
        {
            try
            {
                foreach (var ED in lstEducationInfo)
                    PersonDataProvider.SaveEducationDocument(ED);
            }
            catch (Exception de)
            {
                WinFormsServ.Error("Ошибка сохранения данных об образовании - данные не были сохранены. \n" + de.Message);
            }
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
                Guid? perId = (from per in context.extPerson
                               where per.Barcode == _personBarc
                               select per.Id).FirstOrDefault();

                MainClass.OpenCardPerson(perId.ToString(), null, null);
            }
        }
    }
}