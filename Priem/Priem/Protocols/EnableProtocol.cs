using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Priem
{
    public class EnableProtocol : ProtocolCard
    {
        //конструктор         
        public EnableProtocol(ProtocolList owner, int iFacultyId, int iStudyBasisId, int iStudyFormId)
            : base(owner, iFacultyId, iStudyBasisId, iStudyFormId)
        {
            _type = ProtocolTypes.EnableProtocol;
        }

        public EnableProtocol() : base()
        { 
            //InitializeComponent(); 
        }

        //дополнительная инициализация
        protected override void InitControls()
        {
            sQuery = @"SELECT DISTINCT ed.extAbit.Sum, ed.extPerson.AttestatSeries, ed.extPerson.AttestatNum, ed.extAbit.Id as Id, ed.extAbit.BAckDoc as backdoc, 
             (ed.extAbit.BAckDoc | ed.extAbit.NotEnabled | case when (NOT ed.hlpMinEgeAbiturient.Id IS NULL) then 'true' else 'false' end) as Red, ed.extAbit.RegNum as Рег_Номер, 
             ed.extPerson.FIO as ФИО, 
             ed.extPerson.EducDocument as Документ_об_образовании, 
             ed.extPerson.PassportSeries + ' №' + ed.extPerson.PassportNumber as Паспорт, 
             extAbit.ObrazProgramNameEx + ' ' + (Case when extAbit.ProfileId IS NULL then '' else extAbit.ProfileName end) as Направление, 
             Competition.NAme as Конкурс, extAbit.BackDoc 
             FROM ed.extAbit INNER JOIN ed.extPerson ON ed.extAbit.PersonId = ed.extPerson.Id   
             LEFT JOIN ed.hlpMinEgeAbiturient ON ed.hlpMinEgeAbiturient.Id = ed.extAbit.Id          
             LEFT JOIN ed.Competition ON ed.Competition.Id = ed.extAbit.CompetitionId";

            base.InitControls();

            this.Text = "Протокол о допуске";
            this.chbEnable.Text = "Добавить всех выбранных слева абитуриентов в протокол о допуске";

            this.chbFilter.Text = "Отфильтровать абитуриентов с проверенными данными";
            this.chbFilter.Visible = true;
        }

        protected override void InitAndFillGrids()
        {
            base.InitAndFillGrids();

            string sFilter = string.Empty;
            sFilter = string.Format(" AND ed.extAbit.BackDoc = 0 AND ed.extAbit.NotEnabled=0 AND ed.extAbit.Id NOT IN (SELECT AbiturientId FROM ed.qProtocolHistory WHERE Excluded=0 AND ProtocolId IN (SELECT Id FROM ed.qProtocol WHERE ISOld=0 AND ProtocolTypeId=1 AND FacultyId ={0} AND StudyFormId = {1} AND StudyBasisId = {2}))", 
                _facultyId.ToString(), _studyFormId.ToString(), _studyBasisId.ToString(), MainClass.studyLevelGroupId);

            if (chbFilter.Checked)
                sFilter += " AND ed.extAbit.Checked > 0";

            //сперва общий конкурс (не общ-преим), т.к. чернобыльцы негодуют - льготы есть, а в протокол не попасть
            FillGrid(dgvRight, sQuery, GetWhereClause("ed.extAbit") + sFilter + " AND ed.extAbit.CompetitionId NOT IN (1,2,7,8)/* AND (ed.extPerson.Privileges=0  OR ed.extAbit.CompetitionId IN (5,6))*/", sOrderby);

            //заполнили левый
            if (_id != null)
            {
                sFilter = string.Format(" WHERE ed.extAbit.Id IN (SELECT AbiturientId FROM ed.qProtocolHistory WHERE ProtocolId = '{0}')", _id.ToString());
                FillGrid(dgvLeft, sQuery, sFilter, sOrderby);
            }
            else //новый
            {
                InitGrid(dgvLeft);
            }

            // заполнение льготников, проверенных советниками
            string query = sQuery + GetWhereClause("ed.extAbit") + sFilter + " AND (ed.extAbit.CompetitionId IN (1,8) OR (ed.extPerson.Privileges>0 AND ed.extAbit.CompetitionId IN (2,7))) AND ed.extAbit.Checked>0 ";

            DataSet ds = MainClass.Bdc.GetDataSet(query);

            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                DataGridViewRow r = new DataGridViewRow();
                r.CreateCells(dgvLeft, false, dr["Id"].ToString(), dr["Red"].ToString(), true, dr["backdoc"].ToString(), dr["Рег_номер"].ToString(), dr["ФИО"].ToString(), dr["Sum"].ToString(), dr["Документ_об_образовании"].ToString(), dr["Направление"].ToString(), dr["Конкурс"].ToString(), dr["Паспорт"].ToString());
                if (bool.Parse(dr["Red"].ToString()))
                {
                    r.Cells[5].Style.ForeColor = Color.Red;
                    r.Cells[6].Style.ForeColor = Color.Red;
                }

                r.Cells[5].Style.BackColor = Color.Green;
                r.Cells[6].Style.BackColor = Color.Green;

                dgvLeft.Rows.Add(r);
            }

            //блокируем редактирование кроме первого столбца
            for (int i = 1; i < dgvLeft.ColumnCount; i++)
                dgvLeft.Columns[i].ReadOnly = true;

            dgvLeft.Update();
        }

        //подготовка нужного грида
        protected override void InitGrid(DataGridView dgv)
        {
            base.InitGrid(dgv);
        }
    }
}
