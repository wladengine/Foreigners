using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.Entity.Core.Objects;
using System.Transactions;
using System.Linq;
using PriemLib;

namespace Priem
{
    public class ChangeCompCelProtocol : ProtocolCard
    {
        //конструктор 
        public ChangeCompCelProtocol(ProtocolList owner, int iFacultyId, int iStudyBasisId, int iStudyFormId)
            : this(owner, iFacultyId, iStudyBasisId, iStudyFormId, null)
        {
        }

        //конструктор 
        public ChangeCompCelProtocol(ProtocolList owner, int iFacultyId, int iStudyBasisId, int iStudyFormId, Guid? ProtocolId)
            : base(owner, iFacultyId, iStudyBasisId, iStudyFormId, ProtocolId)
        {
            _type = ProtocolTypes.ChangeCompCelProtocol;
            base.sQuery = this.sQuery;
        }

        //дополнительная инициализация
        protected override void InitControls()
        {
            sQuery = @"SELECT DISTINCT ed.extAbit.Sum, ed.extPerson.AttestatSeries, ed.extPerson.AttestatNum, ed.extAbit.Id as Id, ed.extAbit.BAckDoc as backdoc, 
             (ed.extAbit.BAckDoc | ed.extAbit.NotEnabled) as Red, ed.extAbit.RegNum as Рег_Номер, 
             ed.extPerson.FIO as ФИО, 
             ed.extPerson.EducDocument as Документ_об_образовании, 
             ed.extPerson.PassportSeries + ' №' + ed.extPerson.PassportNumber as Паспорт, 
             extAbit.ObrazProgramNameEx + ' ' + (Case when extAbit.ProfileId IS NULL then '' else extAbit.ProfileName end) as Направление, 
             Competition.NAme as Конкурс, extAbit.BackDoc 
             FROM ed.extAbit INNER JOIN ed.extPerson ON ed.extAbit.PersonId = ed.extPerson.Id                
             LEFT JOIN ed.Competition ON ed.Competition.Id = ed.extAbit.CompetitionId
             INNER JOIN ed.qProtocolHistory ON ed.qProtocolHistory.AbiturientId = ed.extAbit.Id";

            base.InitControls();

            this.Text = "Протокол об изменении типа конкурса целевикам";
            this.chbEnable.Text = "Изменить всем выбранным слева абитуриентам тип конкурса с целевого на дополнительный/общий";

            this.chbFilter.Text = "Отфильтровать абитуриентов с проверенными данными";
            this.chbFilter.Visible = false;
        }

        protected override void InitAndFillGrids()
        {
            base.InitAndFillGrids();

            string sFilter = string.Empty;
            sFilter = " AND ed.extAbit.CompetitionId = 6 ";

            FillGrid(dgvRight, sQuery, GetWhereClause("ed.extAbit") + sFilter, sOrderby);

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
        }

        //подготовка нужного грида
        protected override void InitGrid(DataGridView dgv)
        {
            base.InitGrid(dgv);
        }

        protected override bool Save()
        {
            base.Save();
            try
            {
                ArrayList alQueries = new ArrayList();

                using (PriemEntities context = new PriemEntities())
                {
                    using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                    {
                        try
                        {
                            //меняет тип конкурса
                            foreach (DataGridViewRow r in dgvLeft.Rows)
                            {
                                SortedList slVals = new SortedList();
                                Guid? abitId = new Guid(r.Cells["Id"].Value.ToString());
                                                                

                                int? compNew = (from ab in context.extAbit
                                         where ab.Id == abitId
                                         select ab.OtherCompetitionId).FirstOrDefault();
                                if (compNew == null)
                                    compNew = 4;
                                if (compNew < 1 || compNew > 7)
                                    compNew = 4;

                                context.Abiturient_UpdateCompetititon(compNew, null, false, abitId);
                            }

                            transaction.Complete();
                        }
                        catch (Exception exc)
                        {
                            throw new Exception("Ошибка при сохранении данных: " + exc.Message);
                        }
                    }
                }               
                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при изменении типа конкурса: " + ex.Message);
                return false;
            }
        }
    }
}
