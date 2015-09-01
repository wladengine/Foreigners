using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Transactions;

using BDClassLib;
using EducServLib;
using WordOut;
using PriemLib;

namespace Priem
{
    public partial class DisEntryViewProtocol : ProtocolCard
    {
        public DisEntryViewProtocol(ProtocolList owner, int iStudyLevelGroupId, int sFac, int sSection, int sForm, int? sProf, bool isSec, bool isReduced, bool isParal, bool isList)
            : this(owner, iStudyLevelGroupId, sFac, sSection, sForm, sProf, isSec, isReduced, isParal, isList, null)
        {
        }

        //конструктор 
        public DisEntryViewProtocol(ProtocolList owner, int iStudyLevelGroupId, int sFac, int sSection, int sForm, int? sProf, bool isSec, bool isReduced, bool isParal, bool isList, Guid? sProtocol)
            : base(owner, iStudyLevelGroupId, sFac, sSection, sForm, sProf, isSec, isReduced, isParal, isList, sProtocol)
        {
            _type = ProtocolTypes.DisEntryView;                      
        }

        //дополнительная инициализация
        protected override void InitControls()
        {
            sQuery = string.Format("SELECT DISTINCT extAbitMarksSum.TotalSum as Sum, extAbit.Id as Id, extAbit.BAckDoc as backdoc, " +
            " 'false' as Red, extAbit.RegNum as Рег_Номер, " +
            " extPerson.FIO as ФИО, " +
            " extPerson.EducDocument as Документ_об_образовании, " +
            " extPerson.PassportSeries + ' №' + extPerson.PassportNumber as Паспорт, " +
            " extAbit.ObrazProgramName + ' ' + (Case when extAbit.ProfileName IS NULL then '' else extAbit.ProfileName end) as Направление, " +
            " Competition.NAme as Конкурс, extAbit.BackDoc " +
            " FROM ed.extAbit INNER JOIN ed.extPerson ON extAbit.PersonId = extPerson.Id " +
            " INNER JOIN ed.extEnableProtocol ON extAbit.Id = extEnableProtocol.AbiturientId " +
            " INNER JOIN ed.extEntryView ON extAbit.Id = extEntryView.AbiturientId " +
            " INNER JOIN ed.extAbitMarksSum ON extAbit.Id = extAbitMarksSum.Id " +
            " LEFT JOIN ed.Competition ON Competition.Id = extAbit.CompetitionId ");
          
            string q = string.Format(@"SELECT DISTINCT CONVERT(varchar(100), Id) AS Id, Number as Name FROM ed.extEntryView WHERE IsForeign = 1 AND Excluded=0 AND IsOld=0 
AND FacultyId ={0} AND StudyFormId = {1} AND StudyBasisId = {2} {3} AND IsSecond = {4} AND IsReduced = {5} AND IsParallel = {6} AND IsListener = {7} AND StudyLevelGroupId = {8}", 
_facultyId, _studyFormId, _studyBasisId, 
_licenseProgramId.HasValue ? string.Format("AND LicenseProgramId = {0}", _licenseProgramId) : "", 
QueryServ.StringParseFromBool(_isSecond), 
QueryServ.StringParseFromBool(_isReduced), QueryServ.StringParseFromBool(_isParallel), QueryServ.StringParseFromBool(_isListener), _studyLevelGroupId);
            ComboServ.FillCombo(cbHeaders, HelpClass.GetComboListByQuery(q), false, false);

            cbHeaders.Visible = true;
            lblHeaderText.Text = "Из представления к зачислению №";            
            chbInostr.Visible = true;

            cbHeaders.SelectedIndexChanged += new EventHandler(cbHeaders_SelectedIndexChanged);           
            chbInostr.CheckedChanged += new System.EventHandler(UpdateGrids);  
            
            base.InitControls();

            this.Text = "Приказ об исключении ";
            this.chbEnable.Text = "Добавить всех выбранных слева абитуриентов в приказ об исключении";

            this.chbFilter.Visible = false;
        }

        void cbHeaders_SelectedIndexChanged(object sender, EventArgs e)
        {
            Guid gId = Guid.NewGuid();
            if (Guid.TryParse(HeaderId, out gId))
                _parentProtocolId = gId;

            UpdateGrids();
        }

        public string HeaderId
        {
            get { return ComboServ.GetComboId(cbHeaders); }
            set { ComboServ.SetComboId(cbHeaders, value); }
        }   

        protected override void InitAndFillGrids()
        {
            base.InitAndFillGrids();

            UpdateRight();
            string sFilter = string.Empty;
                        
            //заполнили левый
            if (_id!=null)
            {
                sFilter = string.Format(" WHERE ed.extAbit.Id IN (SELECT AbiturientId FROM ed.qProtocolHistory WHERE ProtocolId = '{0}')", _id);
                FillGrid(dgvLeft, sQuery, sFilter, sOrderby);
            }
            else //новый
            {
                InitGrid(dgvLeft);
            }                    
        }

        private void UpdateGrids(object sender, EventArgs e)
        {
            UpdateGrids();
        }

        void UpdateGrids()
        {
            UpdateRight();
            InitGrid(dgvLeft);
        }

        void UpdateRight()
        {
            if (HeaderId == null)
            {
                dgvRight.DataSource = null;
                return;
            }

            string sFilter = string.Empty;
            sFilter = string.Format(" AND ed.extEntryView.Id = '{0}' ", HeaderId);
            FillGrid(dgvRight, sQuery, GetWhereClause("ed.extAbit") + sFilter, sOrderby);
        }        

        //подготовка нужного грида
        protected override void InitGrid(DataGridView dgv)
        {
            base.InitGrid(dgv);

            dgv.Columns["Pasport"].Visible = false;
            dgv.Columns["Attestat"].Visible = false;
        }
   }
}