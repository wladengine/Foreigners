using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

using EducServLib;

namespace Priem
{
    public partial class CardLoadEntry : Form
    {
        private DBPriem _bdcEduc;

        public CardLoadEntry()
        {
            InitializeComponent();
            InitDB();
        }

        private void InitDB()
        {
            _bdcEduc = new DBPriem();
            try
            {
                _bdcEduc.OpenDatabase(DBConstants.CS_STUDYPLAN);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        private void btnLoadAll_Click(object sender, EventArgs e)
        {
            if (!MainClass.IsEntryChanger())
                return;

            DataSet dsEntry = _bdcEduc.GetDataSet("SELECT * FROM ed.extCurrentEntry");
            DataTable dt = dsEntry.Tables[0];

            using (PriemEntities context = new PriemEntities())
            {
                foreach (DataRow dr in dt.Rows)
                {
                    Guid entryId = (Guid)dr["Id"];

                    int cntEnt = (from ent in context.Entry
                                  where ent.Id == entryId
                                  select ent).Count();

                    if (cntEnt > 0)
                        continue;

                    Entry item = new Entry();
                    item.Id = entryId;
                    item.FacultyId = (int)dr["FacultyId"];
                    item.LicenseProgramId = (int)dr["LicenseProgramId"];
                    item.LicenseProgramName = dr["LicenseProgramName"].ToString();
                    item.LicenseProgramCode = dr["LicenseProgramCode"].ToString();
                    item.ObrazProgramId = (int)dr["ObrazProgramId"];
                    item.ObrazProgramName = dr["ObrazProgramName"].ToString();
                    item.ObrazProgramNumber = dr["ObrazProgramNumber"].ToString();
                    item.ObrazProgramCrypt = dr["ObrazProgramCrypt"].ToString();
                    item.ProfileId = dr.Field<Guid?>("ProfileId"); 
                    item.ProfileName = dr["ProfileName"].ToString();
                    item.StudyBasisId = (int)dr["StudyBasisId"];
                    item.StudyFormId = (int)dr["StudyFormId"];
                    item.StudyLevelId = (int)dr["StudyLevelId"];
                    item.StudyPlanId = (Guid)dr["StudyPlanId"];
                    item.StudyPlanNumber = dr["StudyPlanNumber"].ToString();
                    item.ProgramModeShortName = dr["ProgramModeShortName"].ToString();
                    item.IsSecond = (bool)dr["IsSecond"];                    
                    item.KCP = dr.Field<int?>("KCP");
                    
                    context.Entry_Insert(entryId, (int)dr["FacultyId"], (int)dr["LicenseProgramId"], dr["LicenseProgramName"].ToString(), 
                            dr["LicenseProgramCode"].ToString(), (int)dr["ObrazProgramId"], dr["ObrazProgramName"].ToString(), dr["ObrazProgramNumber"].ToString(),
                            dr["ObrazProgramCrypt"].ToString(), dr.Field<Guid?>("ProfileId"), dr["ProfileName"].ToString(), (int)dr["StudyBasisId"],
                            (int)dr["StudyFormId"], (int)dr["StudyLevelId"], (Guid)dr["StudyPlanId"], dr["StudyPlanNumber"].ToString(),
                            dr["ProgramModeShortName"].ToString(), (bool)dr["IsSecond"], (bool)dr["IsReduced"], (bool)dr["IsParallel"], dr.Field<int?>("KCP"));

                }

                MessageBox.Show("Выполнено");
            }
        }   

        private void CardLoadEntry_FormClosing(object sender, FormClosingEventArgs e)
        {
            _bdcEduc.CloseDataBase();
        }

        private void btnCheckUpdates_Click(object sender, EventArgs e)
        {
            if (!MainClass.IsEntryChanger())
                return;
            
            List<string> missingInOurs = new List<string>();
            List<string> invalidOurs = new List<string>();
            List<string> extraOurs = new List<string>();
                        
            DataSet ds = _bdcEduc.GetDataSet(string.Format("SELECT * FROM ed.extCurrentEntry"));
            string diskDriveLetter = System.Environment.UserName == "o.belenog" ? "O" : "D";
            using (StreamWriter sw = new StreamWriter(string.Format("{0}:\\result.txt", diskDriveLetter)))
            {
                List<string> lstOld = new List<string>();
                
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    string entryId = dr["Id"].ToString();

                    lstOld.Add(string.Format("'{0}'", entryId));

                    DataSet dsOur = MainClass.Bdc.GetDataSet(string.Format("SELECT * FROM ed.Entry WHERE Id = '{0}'", entryId));
                    if (dsOur.Tables[0].Rows.Count == 0)
                    {
                        missingInOurs.Add(entryId);
                        continue;
                    }

                    DataRow drOur = dsOur.Tables[0].Rows[0];

                    if (!dr["FacultyId"].Equals(drOur["FacultyId"]))
                        invalidOurs.Add(entryId + ": FacultyId Old - " + drOur["FacultyId"].ToString() + "; New - " + dr["FacultyId"].ToString());

                    if (!dr["LicenseProgramId"].Equals(drOur["LicenseProgramId"]))
                        invalidOurs.Add(entryId + ": LicenseProgramId Old - " + drOur["LicenseProgramId"].ToString() + "; New - " + dr["LicenseProgramId"].ToString());

                    if (!dr["LicenseProgramName"].Equals(drOur["LicenseProgramName"]))
                        invalidOurs.Add(entryId + ": LicenseProgramName Old - " + drOur["LicenseProgramName"].ToString() + "; New - " + dr["LicenseProgramName"].ToString());

                    if (!dr["LicenseProgramCode"].Equals(drOur["LicenseProgramCode"]))
                        invalidOurs.Add(entryId + ": LicenseProgramCode Old - " + drOur["LicenseProgramCode"].ToString() + "; New - " + dr["LicenseProgramCode"].ToString());

                    if (!dr["ObrazProgramId"].Equals(drOur["ObrazProgramId"]))
                        invalidOurs.Add(entryId + ": ObrazProgramId Old - " + drOur["ObrazProgramId"].ToString() + "; New - " + dr["ObrazProgramId"].ToString());

                    if (!dr["ObrazProgramName"].Equals(drOur["ObrazProgramName"]))
                        invalidOurs.Add(entryId + ": ObrazProgramName Old - " + drOur["ObrazProgramName"].ToString() + "; New - " + dr["ObrazProgramName"].ToString());

                    if (!dr["ObrazProgramNumber"].Equals(drOur["ObrazProgramNumber"]))
                        invalidOurs.Add(entryId + ": ObrazProgramNumber Old - " + drOur["ObrazProgramNumber"].ToString() + "; New - " + dr["FacultyId"].ToString());

                    if (!dr["ObrazProgramCrypt"].Equals(drOur["ObrazProgramCrypt"]))
                        invalidOurs.Add(entryId + ": ObrazProgramCrypt Old - " + drOur["ObrazProgramCrypt"].ToString() + "; New - " + dr["ObrazProgramCrypt"].ToString());

                    if (!dr["ProfileId"].Equals(drOur["ProfileId"]))
                        invalidOurs.Add(entryId + ": ProfileId Old - " + drOur["ProfileId"].ToString() + "; New - " + dr["ProfileId"].ToString());

                    if (!dr["ProfileName"].Equals(drOur["ProfileName"]))
                        invalidOurs.Add(entryId + ": ProfileName Old - " + drOur["ProfileName"].ToString() + "; New - " + dr["ProfileName"].ToString());

                    if (!dr["StudyBasisId"].Equals(drOur["StudyBasisId"]))
                        invalidOurs.Add(entryId + ": StudyBasisId Old - " + drOur["StudyBasisId"].ToString() + "; New - " + dr["StudyBasisId"].ToString());

                    if (!dr["StudyFormId"].Equals(drOur["StudyFormId"]))
                        invalidOurs.Add(entryId + ": StudyFormId Old - " + drOur["StudyFormId"].ToString() + "; New - " + dr["StudyFormId"].ToString());

                    if (!dr["StudyLevelId"].Equals(drOur["StudyLevelId"]))
                        invalidOurs.Add(entryId + ": StudyLevelId Old - " + drOur["StudyLevelId"].ToString() + "; New - " + dr["StudyLevelId"].ToString());

                    if (!dr["StudyPlanId"].Equals(drOur["StudyPlanId"]))
                        invalidOurs.Add(entryId + ": StudyPlanId Old - " + drOur["StudyPlanId"].ToString() + "; New - " + dr["StudyPlanId"].ToString());

                    if (!dr["StudyPlanNumber"].Equals(drOur["StudyPlanNumber"]))
                        invalidOurs.Add(entryId + ": StudyPlanNumber Old - " + drOur["StudyPlanNumber"].ToString() + "; New - " + dr["StudyPlanNumber"].ToString());

                    if (!dr["ProgramModeShortName"].Equals(drOur["ProgramModeShortName"]))
                        invalidOurs.Add(entryId + ": ProgramModeShortName Old - " + drOur["ProgramModeShortName"].ToString() + "; New - " + dr["ProgramModeShortName"].ToString());

                    if (!dr["IsSecond"].Equals(drOur["IsSecond"]))
                        invalidOurs.Add(entryId + ": IsSecond Old - " + drOur["IsSecond"].ToString() + "; New - " + dr["IsSecond"].ToString());

                    if (!dr["IsReduced"].Equals(drOur["IsReduced"]))
                        invalidOurs.Add(entryId + ": IsReduced Old - " + drOur["IsReduced"].ToString() + "; New - " + dr["IsReduced"].ToString());

                    if (!dr["IsParallel"].Equals(drOur["IsParallel"]))
                        invalidOurs.Add(entryId + ": IsParallel Old - " + drOur["IsParallel"].ToString() + "; New - " + dr["IsParallel"].ToString());

                    if (!dr["KCP"].Equals(drOur["KCP"]))
                        invalidOurs.Add(entryId + ": KCP Old - " + drOur["KCP"].ToString() + "; New - " + dr["KCP"].ToString());
                }

                DataSet dsExtra = MainClass.Bdc.GetDataSet(string.Format("SELECT ed.Entry.Id FROM ed.Entry WHERE Id NOT IN ({0})", Util.BuildStringWithCollection(lstOld)));
                foreach (DataRow dr in dsExtra.Tables[0].Rows)
                {
                    extraOurs.Add(dr["Id"].ToString());
                }
                
                sw.WriteLine("Лишние нас:");
                sw.WriteLine("");
                foreach (string pl in extraOurs)
                {
                    sw.WriteLine(pl);
                }

                sw.WriteLine("Отсутсвуют у нас:");
                sw.WriteLine("");
                foreach (string pl in missingInOurs)
                {
                    sw.WriteLine(pl);
                }

                sw.WriteLine("");
                sw.WriteLine("Другие значение у нас:");
                sw.WriteLine("");

                foreach (string pl in invalidOurs)
                {
                    sw.WriteLine(pl);
                }

                MessageBox.Show("Выполнено");

                Process pr = new Process();
                pr.StartInfo.Verb = "Open";
                pr.StartInfo.FileName = string.Format("{0}:\\result.txt", diskDriveLetter);
                pr.Start();
            }            
        }

        private void btnLoadUpdates_Click(object sender, EventArgs e)
        {
            if (!MainClass.IsEntryChanger())
                return;           
            
            DataSet ds = _bdcEduc.GetDataSet(string.Format("SELECT * FROM ed.extCurrentEntry"));

            using (PriemEntities context = new PriemEntities())
            {
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    Guid entryId = (Guid)dr["Id"];

                    int cntEnt = (from ent in context.Entry
                                  where ent.Id == entryId
                                  select ent).Count();

                    if (cntEnt == 0)
                    {
                        Entry item = new Entry();
                        item.Id = entryId;
                        item.FacultyId = (int)dr["FacultyId"];
                        item.LicenseProgramId = (int)dr["LicenseProgramId"];
                        item.LicenseProgramName = dr["LicenseProgramName"].ToString();
                        item.LicenseProgramCode = dr["LicenseProgramCode"].ToString();
                        item.ObrazProgramId = (int)dr["ObrazProgramId"];
                        item.ObrazProgramName = dr["ObrazProgramName"].ToString();
                        item.ObrazProgramNumber = dr["ObrazProgramNumber"].ToString();
                        item.ObrazProgramCrypt = dr["ObrazProgramCrypt"].ToString();
                        item.ProfileId = dr.Field<Guid?>("ProfileId");
                        item.ProfileName = dr["ProfileName"].ToString();
                        item.StudyBasisId = (int)dr["StudyBasisId"];
                        item.StudyFormId = (int)dr["StudyFormId"];
                        item.StudyLevelId = (int)dr["StudyLevelId"];
                        item.StudyPlanId = (Guid)dr["StudyPlanId"];
                        item.StudyPlanNumber = dr["StudyPlanNumber"].ToString();
                        item.ProgramModeShortName = dr["ProgramModeShortName"].ToString();
                        item.IsSecond = (bool)dr["IsSecond"];
                        item.KCP = dr.Field<int?>("KCP");
                        
                        context.Entry_Insert(entryId, (int)dr["FacultyId"], (int)dr["LicenseProgramId"], dr["LicenseProgramName"].ToString(),
                                dr["LicenseProgramCode"].ToString(), (int)dr["ObrazProgramId"], dr["ObrazProgramName"].ToString(), dr["ObrazProgramNumber"].ToString(),
                                dr["ObrazProgramCrypt"].ToString(), dr.Field<Guid?>("ProfileId"), dr["ProfileName"].ToString(), (int)dr["StudyBasisId"],
                                (int)dr["StudyFormId"], (int)dr["StudyLevelId"], (Guid)dr["StudyPlanId"], dr["StudyPlanNumber"].ToString(),
                                dr["ProgramModeShortName"].ToString(), (bool)dr["IsSecond"], (bool)dr["IsReduced"],(bool)dr["IsParallel"], dr.Field<int?>("KCP"));
                    }
                }

                MessageBox.Show("Выполнено");
            }            
        }

        private void btnUpdateKCP_Click(object sender, EventArgs e)
        {
            if (!MainClass.IsEntryChanger())
                return;

            DataSet ds = _bdcEduc.GetDataSet(string.Format("SELECT * FROM ed.extCurrentEntry"));
            using (PriemEntities context = new PriemEntities())
            {
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    Guid entryId = (Guid)dr["Id"];

                    int cntEnt = (from ent in context.Entry
                                  where ent.Id == entryId
                                  select ent).Count();
                   
                   if (cntEnt == 0) 
                       continue;  

                    Entry entry =  (from ent in context.Entry
                                    where ent.Id == entryId
                                    select ent).FirstOrDefault();

                    int? kcpSP;
                    
                    if (dr["KCP"].ToString() == string.Empty)
                        kcpSP = 0;
                    else
                        kcpSP = (int?)dr["KCP"];

                    if (kcpSP != entry.KCP)
                        context.Entry_UpdateKC(entryId, kcpSP);
                    
                }                
                MessageBox.Show("Выполнено");                
            }            
        }           
    }
}
