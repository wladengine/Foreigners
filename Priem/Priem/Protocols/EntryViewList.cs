using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using BaseFormsLib;
using EducServLib;
using WordOut;
using PriemLib;

namespace Priem
{
    public partial class EntryViewList : PriemLib.EntryViewList
    {
        //private DBPriem _bdc;
        //private string sQuery;
        protected ProtocolRefreshHandler prh = null;

        public EntryViewList()
            :base()
        {            
            //this.CenterToParent();
            //this.MdiParent = MainClass.mainform;

            //this.sQuery = string.Format("SELECT DISTINCT Person.Id, {0} as Ид_номер, Person.Surname AS Фамилия, Person.Name AS Имя, Person.SecondName AS Отчество, Person.BirthDate AS Дата_рождения " +
            //                       "FROM Person INNER JOIN ExamsVedHistory ON ExamsVedHistory.PersonId = Person.Id ", MainClass.GetStringPersonNumber());

            //InitializeComponent();
            //InitControls();
            
            //btnPrintOrder.Visible = btnPrintOrder.Enabled = btnCancelView.Enabled = btnCancelView.Visible = false;
            btnCreate.Enabled = false;

            chbIsForeign.Enabled = false;
            chbIsForeign.Checked = true;

            if (MainClass.IsFacMain())
            {
                btnCreate.Enabled = true;
                btnPrintOrder.Visible = btnPrintOrder.Enabled = true;
            }

            if (MainClass.IsPasha())
            {                
                btnPrintOrder.Visible = btnPrintOrder.Enabled = btnCancelView.Enabled = btnCancelView.Visible = true;
                btnCreate.Enabled = true;
            }

            if (MainClass.IsPrintOrder())
            {
                btnPrintOrder.Visible = btnPrintOrder.Enabled = true;
            }

            //// посомтреть, почему отдельные факультеты
            //if (_bdc.IsMed() || _bdc.GetFacultyId() == "9" || _bdc.GetFacultyId() == "14" || _bdc.GetFacultyId() == "20")
            //    btnCreate.Enabled = true;

            //if (_bdc.IsReadOnly())
            //    btnPrintOrder.Visible = btnPrintOrder.Enabled = true;
        }

        //дополнительная инициализация контролов
        public override void InitControls()
        {
            InitFocusHandlers();
            _bdc = MainClass.Bdc;

            ComboServ.FillCombo(cbStudyLevelGroup, HelpClass.GetComboListByTable("ed.StudyLevelGroup", "ORDER BY Acronym"), false, false);
            ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, false);

            cbStudyBasis.SelectedIndex = 0;
            FillFaculty();
            FillStudyForm();
            FillLicenseProgram();             

            UpdateDataGrid();

            cbStudyLevelGroup.SelectedIndexChanged += cbStudyLevel_SelectedIndexChanged;
            cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);
            cbStudyForm.SelectedIndexChanged += new EventHandler(cbStudyForm_SelectedIndexChanged);
            cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);
            cbLicenseProgram.SelectedIndexChanged += new EventHandler(cbLicenseProgram_SelectedIndexChanged);              

            prh = new ProtocolRefreshHandler(UpdateDataGrid);
            MainClass.AddProtocolHandler(prh);          
        }

        void cbStudyLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillFaculty();
        }

        //void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    FillStudyForm();            
        //}
 
        //void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    FillStudyForm();            
        //}
        //void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    FillLicenseProgram();
        //}
        //void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    UpdateDataGrid();
        //}

        //public int? FacultyId
        //{
        //    get { return ComboServ.GetComboIdInt(cbFaculty); }
        //    set { ComboServ.SetComboId(cbFaculty, value); }
        //}
        //public int? LicenseProgramId
        //{
        //    get { return ComboServ.GetComboIdInt(cbLicenseProgram); }
        //    set { ComboServ.SetComboId(cbLicenseProgram, value); }
        //}
        //public int? StudyBasisId
        //{
        //    get { return ComboServ.GetComboIdInt(cbStudyBasis); }
        //    set { ComboServ.SetComboId(cbStudyBasis, value); }
        //}
        //public int? StudyFormId
        //{
        //    get { return ComboServ.GetComboIdInt(cbStudyForm); }
        //    set { ComboServ.SetComboId(cbStudyForm, value); }
        //}
        //public int StudyLevelGroupId
        //{
        //    get { return ComboServ.GetComboIdInt(cbStudyLevelGroup).Value; }
        //    set { ComboServ.SetComboId(cbStudyLevelGroup, value); }
        //}

        //public bool IsSecond
        //{
        //    get { return chbIsSecond.Checked; }
        //    set { chbIsSecond.Checked = value; }
        //}
        //public bool IsReduced
        //{
        //    get { return chbIsReduced.Checked; }
        //    set { chbIsReduced.Checked = value; }
        //}
        //public bool IsParallel
        //{
        //    get { return chbIsParallel.Checked; }
        //    set { chbIsParallel.Checked = value; }
        //}
        //public bool IsListener
        //{
        //    get { return chbIsListener.Checked; }
        //    set { chbIsListener.Checked = value; }
        //}

        private void FillFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.StudyLevelGroupId == StudyLevelGroupId);
                List<KeyValuePair<string, string>> lst = ent.Select(x => new { x.FacultyName, x.FacultyId }).Distinct().ToList()
                    .Select(u => new KeyValuePair<string, string>(u.FacultyId.ToString(), u.FacultyName)).OrderBy(x => x.Value).Distinct().ToList();

                ComboServ.FillCombo(cbFaculty, lst, false, false);
            }
        }
        //private void FillStudyForm()
        //{
        //    using (PriemEntities context = new PriemEntities())
        //    {
        //        var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId).Where(c => c.StudyBasisId == StudyBasisId);

        //        ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

        //        List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.StudyFormId.ToString(), u.StudyFormName)).Distinct().ToList();

        //        ComboServ.FillCombo(cbStudyForm, lst, false, false);
        //    }
        //}
        public override void  FillLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                if (FacultyId != null)
                {
                    var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                    ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

                    if (StudyBasisId != null)
                        ent = ent.Where(c => c.StudyBasisId == StudyBasisId);
                    if (StudyFormId != null)
                        ent = ent.Where(c => c.StudyFormId == StudyFormId);

                    List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.LicenseProgramId.ToString(), u.LicenseProgramCode + " " + u.LicenseProgramName)).Distinct().ToList();
                    ComboServ.FillCombo(base.cbLicenseProgram, lst, false, true);
                }
                else
                    ComboServ.FillCombo(base.cbLicenseProgram, new List<KeyValuePair<string, string>>(), false, true);
            }
        }

        public override void UpdateDataGrid()
        {
            if (StudyFormId == null || StudyBasisId == null)
            {
                dgvViews.DataSource = null;
                return;
            }
            string qquery = @"SELECT DISTINCT Id, Number + ' (' + CONVERT(nvarchar, Date, 104) + ')' AS 'Номер представления' 
FROM ed.extEntryView 
WHERE StudyFormId={0} AND StudyBasisId={1} "+(FacultyId.HasValue? "AND FacultyId= {2} ":"")+(LicenseProgramId.HasValue ? "AND LicenseProgramId = {3} ":"")+@"AND IsListener = {4} AND IsSecond = {5} 
AND IsReduced = {6} AND IsParallel = {7} AND IsForeign = {8} AND StudyLevelGroupId = {9}
order by 2";
            string query = string.Format(qquery,
                               StudyFormId,
                               StudyBasisId,
                               FacultyId,
                               LicenseProgramId,
                               QueryServ.StringParseFromBool(IsListener),
                               QueryServ.StringParseFromBool(IsSecond),
                               QueryServ.StringParseFromBool(IsReduced),
                               QueryServ.StringParseFromBool(IsParallel),
                               QueryServ.StringParseFromBool(MainClass.dbType == PriemType.PriemForeigners),
                               StudyLevelGroupId);
            HelpClass.FillDataGrid(dgvViews, _bdc, query, "");

            if (dgvViews.Columns.Contains("Номер представления"))
                dgvViews.Columns["Номер представления"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            if (dgvViews.Rows.Count == 0)
            {
                dgvViews.CurrentCell = null;
                ViewProtocolInfo();
            }
        }
        //private void btnCreate_Click(object sender, EventArgs e)
        //{
        //    new EntryViewProtocol(null, StudyLevelGroupId, FacultyId.Value, StudyBasisId.Value, StudyFormId.Value, LicenseProgramId, IsSecond, IsReduced, IsParallel, IsListener, IsCel).Show();            
        //}
        //private void btnPrint_Click(object sender, EventArgs e)
        //{  
        //    if (dgvViews.CurrentRow == null)
        //        return;

        //    if (dgvViews.CurrentRow.Index < 0)
        //        return;

        //    SaveFileDialog sfd = new SaveFileDialog();
        //    sfd.Filter = "ADOBE Pdf files|*.pdf";
        //    if (sfd.ShowDialog() == DialogResult.OK)
        //        Print.PrintEntryView(dgvViews.CurrentRow.Cells["Id"].Value.ToString(), sfd.FileName);
        //}
        //private void btnPrintOrder_Click(object sender, EventArgs e)
        //{
        //    if (dgvViews.CurrentRow == null)
        //        return;

        //    if (dgvViews.CurrentRow.Index < 0)
        //        return;

        //    Guid protocolId = (Guid)dgvViews.CurrentRow.Cells["Id"].Value;

        //    Print.PrintOrder(protocolId, chbCel.Checked);
        //}
        //private void btnOrderReview_Click(object sender, EventArgs e)
        //{
        //    if (dgvViews.CurrentRow == null)
        //        return;

        //    if (dgvViews.CurrentRow.Index < 0)
        //        return;

        //    Guid protocolId = (Guid)dgvViews.CurrentRow.Cells["Id"].Value;

        //    Print.PrintOrderReview(protocolId);
        //}
        //private void btnCancelView_Click(object sender, EventArgs e)
        //{
        //    if(!MainClass.IsPasha())
        //        return;

        //    using(PriemEntities context = new PriemEntities())
        //    {
        //        if (dgvViews.CurrentRow == null || dgvViews.CurrentRow.Index < 0)
        //            return;

        //        if (MessageBox.Show("Отменить выделенное представление", "Внимание!", MessageBoxButtons.YesNo) == DialogResult.Yes)
        //        {
        //            Guid? protocolId = new Guid(dgvViews.CurrentRow.Cells["Id"].Value.ToString());
        //            context.Protocol_UpdateIsOld(true, protocolId);                   
        //            MessageBox.Show("Представление отменено");
        //            UpdateDataGrid();
        //        }
        //    }
        //}

        //private void EntryViewList_FormClosing(object sender, FormClosingEventArgs e)
        //{
        //    MainClass.RemoveProtocolHandler(prh);
        //}

        //private void chbIsListener_CheckedChanged(object sender, EventArgs e)
        //{
        //    UpdateDataGrid();
        //}
        //private void chbIsSecond_CheckedChanged(object sender, EventArgs e)
        //{
        //    FillStudyForm();
        //}
        //private void chbIsReduced_CheckedChanged(object sender, EventArgs e)
        //{
        //    FillStudyForm();
        //}
        //private void chbIsParallel_CheckedChanged(object sender, EventArgs e)
        //{
        //    FillStudyForm();
        //}     
        public override void ViewProtocolInfo()  
        {
            base.ViewProtocolInfo();
            lblHasForeigners.Visible = false;
        }
    }
}
