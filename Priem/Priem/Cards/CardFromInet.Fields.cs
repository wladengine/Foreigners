using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using EducServLib;

namespace Priem
{
    public partial class CardFromInet
    {
        public string PersonName
        {
            get { return tbName.Text.Trim(); }
            set { tbName.Text = value; }
        }
        public string SecondName
        {
            get { return tbSecondName.Text.Trim(); }
            set { tbSecondName.Text = value; }
        }
        public string Surname
        {
            get { return tbSurname.Text.Trim(); }
            set { tbSurname.Text = value; }
        }
        public DateTime? BirthDate
        {
            get { return dtBirthDate.Value.Date; }
            set
            {
                try
                {
                    if (value.HasValue)
                        dtBirthDate.Value = value.Value;
                }
                catch (ArgumentOutOfRangeException)
                {
                    WinFormsServ.Error("Некорректная дата рождения!");
                }
            }
        }
        public string BirthPlace
        {
            get { return tbBirthPlace.Text.Trim(); }
            set { tbBirthPlace.Text = value; }
        }
        public bool Sex
        {
            get { return rbMale.Checked; }
            set
            {
                rbMale.Checked = value;
                rbFemale.Checked = !value;
            }
        }

        protected int? PassportTypeId
        {
            get { return ComboServ.GetComboIdInt(cbPassportType); }
            set { ComboServ.SetComboId(cbPassportType, value); }
        }
        public string PassportSeries
        {
            get { return tbPassportSeries.Text.Replace(" ", "").Trim(); }
            set { tbPassportSeries.Text = value; }
        }
        public string PassportNumber
        {
            get { return tbPassportNumber.Text.Replace(" ", "").Trim(); }
            set { tbPassportNumber.Text = value; }
        }
        public string PassportAuthor
        {
            get { return tbPassportAuthor.Text.Trim(); }
            set { tbPassportAuthor.Text = value; }
        }
        public DateTime? PassportDate
        {
            get { return dtPassportDate.Value.Date; }
            set
            {
                try
                {
                    if (value.HasValue)
                        dtPassportDate.Value = value.Value;
                }
                catch (ArgumentOutOfRangeException)
                {
                    WinFormsServ.Error("Некорректная дата выдачи паспорта");
                }
            }
        }
        public string SNILS
        {
            get { return tbSNILS.Text.Trim(); }
            set { tbSNILS.Text = value; }
        }
        public string PassportCode
        {
            get { return tbPassportCode.Text.Trim(); }
            set { tbPassportCode.Text = value; }
        }
        public string PersonalCode
        {
            get { return tbPersonalCode.Text.Trim(); }
            set { tbPersonalCode.Text = value; }
        }

        protected int? CountryId
        {
            get { return ComboServ.GetComboIdInt(cbCountry); }
            set { ComboServ.SetComboId(cbCountry, value); }
        }
        protected int? NationalityId
        {
            get { return ComboServ.GetComboIdInt(cbNationality); }
            set { ComboServ.SetComboId(cbNationality, value); }
        }
        protected int? RegionId
        {
            get { return ComboServ.GetComboIdInt(cbRegion); }
            set { ComboServ.SetComboId(cbRegion, value); }
        }
        public string Phone
        {
            get { return tbPhone.Text.Trim(); }
            set { tbPhone.Text = value; }
        }
        public string Mobiles
        {
            get { return tbMobiles.Text.Trim(); }
            set { tbMobiles.Text = value; }
        }
        public string Email
        {
            get { return tbEmail.Text.Trim(); }
            set { tbEmail.Text = value; }
        }
        public string Code
        {
            get { return tbCode.Text.Trim(); }
            set { tbCode.Text = value; }
        }
        public string City
        {
            get { return tbCity.Text.Trim(); }
            set { tbCity.Text = value; }
        }
        public string Street
        {
            get { return tbStreet.Text.Trim(); }
            set { tbStreet.Text = value; }
        }
        public string House
        {
            get { return tbHouse.Text.Trim(); }
            set { tbHouse.Text = value; }
        }
        public string Korpus
        {
            get { return tbKorpus.Text.Trim(); }
            set { tbKorpus.Text = value; }
        }
        public string Flat
        {
            get { return tbFlat.Text.Trim(); }
            set { tbFlat.Text = value; }
        }
        public string CodeReal
        {
            get { return tbCodeReal.Text.Trim(); }
            set { tbCodeReal.Text = value; }
        }
        public string CityReal
        {
            get { return tbCityReal.Text.Trim(); }
            set { tbCityReal.Text = value; }
        }
        public string StreetReal
        {
            get { return tbStreetReal.Text.Trim(); }
            set { tbStreetReal.Text = value; }
        }
        public string HouseReal
        {
            get { return tbHouseReal.Text.Trim(); }
            set { tbHouseReal.Text = value; }
        }
        public string KorpusReal
        {
            get { return tbKorpusReal.Text.Trim(); }
            set { tbKorpusReal.Text = value; }
        }
        public string FlatReal
        {
            get { return tbFlatReal.Text.Trim(); }
            set { tbFlatReal.Text = value; }
        }
        public string KladrCode
        {
            get { return tbKladrCode.Text.Trim(); }
            set { tbKladrCode.Text = value; }
        }
        public bool HostelAbit
        {
            get { return chbHostelAbitYes.Checked; }
            set
            {
                chbHostelAbitYes.Checked = value;
                chbHostelAbitNo.Checked = !value;
            }
        }
        public bool HostelEduc
        {
            get { return chbHostelEducYes.Checked; }
            set
            {
                chbHostelEducYes.Checked = value;
                chbHostelEducNo.Checked = !value;
            }
        }

        protected int? LanguageId
        {
            get { return ComboServ.GetComboIdInt(cbLanguage); }
            set { ComboServ.SetComboId(cbLanguage, value); }
        }
        public string SchoolCity
        {
            get { return cbSchoolCity.Text.Trim(); }
            set { cbSchoolCity.Text = value; }
        }
        protected int? SchoolTypeId
        {
            get { return ComboServ.GetComboIdInt(cbSchoolType); }
            set { ComboServ.SetComboId(cbSchoolType, value); }
        }
        public string SchoolName
        {
            get { return tbSchoolName.Text.Trim(); }
            set { tbSchoolName.Text = value; }
        }
        public string SchoolNum
        {
            get { return tbSchoolNum.Text.Trim(); }
            set { tbSchoolNum.Text = value; }
        }
        public int? SchoolExitYear
        {
            get
            {
                int j;
                if (int.TryParse(tbSchoolExitYear.Text.Trim(), out j))
                    return j;
                else
                    return null;
            }
            set { tbSchoolExitYear.Text = Util.ToStr(value); }
        }
        protected int? CountryEducId
        {
            get { return ComboServ.GetComboIdInt(cbCountryEduc); }
            set { ComboServ.SetComboId(cbCountryEduc, value); }
        }
        protected int? RegionEducId
        {
            get { return ComboServ.GetComboIdInt(cbRegionEduc); }
            set { ComboServ.SetComboId(cbRegionEduc, value); }
        }
        public bool IsEqual
        {
            get { return chbEkvivEduc.Checked; }
            set { chbEkvivEduc.Checked = value; }
        }
        public string EqualDocumentNumber
        {
            get { return tbEqualityDocumentNumber.Text.Trim(); }
            set { tbEqualityDocumentNumber.Text = value; }
        }
        public bool IsExcellent
        {
            get { return chbIsExcellent.Checked; }
            set { chbIsExcellent.Checked = value; }
        }
        public string AttestatSeries
        {
            get { return cbAttestatSeries.Text.Trim(); }
            set { cbAttestatSeries.Text = value; }
        }
        public string AttestatNum
        {
            get { return tbAttestatNum.Text.Trim(); }
            set { tbAttestatNum.Text = value; }
        }
        public string DiplomSeries
        {
            get { return tbDiplomSeries.Text.Trim(); }
            set { tbDiplomSeries.Text = value; }
        }
        public string DiplomNum
        {
            get { return tbDiplomNum.Text.Trim(); }
            set { tbDiplomNum.Text = value; }
        }
        public string HighEducation
        {
            get { return tbSchoolName.Text.Trim(); }
            set { tbSchoolName.Text = value; }
        }
        public string HEProfession
        {
            get { return tbHEProfession.Text.Trim(); }
            set { tbHEProfession.Text = value; }
        }
        public string HEQualification
        {
            get { return cbHEQualification.Text.Trim(); }
            set 
            {
                if (cbHEQualification.Items.Contains(value))
                    cbHEQualification.SelectedItem = value;
                else
                    cbHEQualification.Text = value;
            }
        }
        public int? HEEntryYear
        {
            get
            {
                int j;
                if (int.TryParse(tbHEEntryYear.Text.Trim(), out j))
                    return j;
                else
                    return null;
            }
            set { tbHEEntryYear.Text = Util.ToStr(value); }
        }
        public int? HEExitYear
        {
            get
            {
                int j;
                if (int.TryParse(tbHEExitYear.Text.Trim(), out j))
                    return j;
                else
                    return null;
            }
            set { tbHEExitYear.Text = Util.ToStr(value); }
        }
        public string HEWork
        {
            get { return tbHEWork.Text.Trim(); }
            set { tbHEWork.Text = value; }
        }
        protected int? HEStudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbHEStudyForm); }
            set { ComboServ.SetComboId(cbHEStudyForm, value); }
        }
        public string Stag
        {
            get { return tbStag.Text.Trim(); }
            set { tbStag.Text = value; }
        }
        public string WorkPlace
        {
            get { return tbWorkPlace.Text.Trim(); }
            set { tbWorkPlace.Text = value; }
        }
        public string MSVuz
        {
            get { return tbMSVuz.Text.Trim(); }
            set { tbMSVuz.Text = value; }
        }
        public string MSCourse
        {
            get { return tbMSCourse.Text.Trim(); }
            set { tbMSCourse.Text = value; }
        }
        protected int? MSStudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbMSStudyForm); }
            set { ComboServ.SetComboId(cbMSStudyForm, value); }
        }
        public double? SchoolAVG
        {
            get
            {
                double j;
                if (double.TryParse(tbSchoolAVG.Text.Trim(), out j))
                    return j;
                else
                    return null;
            }
            set
            { tbSchoolAVG.Text = Util.ToStr(value); }
        }

        public string ScienceWork
        {
            get { return tbScienceWork.Text.Trim(); }
            set { tbScienceWork.Text = value; }
        }
        public string ExtraInfo
        {
            get { return tbExtraInfo.Text.Trim(); }
            set { tbExtraInfo.Text = value; }
        }
        public string PersonInfo
        {
            get { return tbPersonInfo.Text.Trim(); }
            set { tbPersonInfo.Text = value; }
        }
        public int? Privileges
        {
            get
            {
                int val = 0;

                if (chbSir.Checked)
                    val += 1;
                if (chbCher.Checked)
                    val += 2;
                if (chbVoen.Checked)
                    val += 4;
                if (chbPSir.Checked)
                    val += 16;
                if (chbInv.Checked)
                    val += 32;
                if (chbBoev.Checked)
                    val += 64;
                if (chbStag.Checked)
                    val += 128;
                if (chbRebSir.Checked)
                    val += 256;
                if (chbExtPoss.Checked)
                    val += 512;

                return val;
            }
            set
            {
                int iPriv;
                if (!value.HasValue)
                    iPriv = 0;
                else
                    iPriv = value.Value;

                int[] masks = { 1, 2, 4, 16, 32, 64, 128, 256, 512 };

                bool[] res = new bool[9];
                for (int i = 0; i < 9; i++)
                    res[i] = (iPriv & masks[i]) != 0;

                chbSir.Checked = res[0];
                chbCher.Checked = res[1];
                chbVoen.Checked = res[2];
                chbPSir.Checked = res[3];
                chbInv.Checked = res[4];
                chbBoev.Checked = res[5];
                chbStag.Checked = res[6];
                chbRebSir.Checked = res[7];
                chbExtPoss.Checked = res[8];
            }
        }
        public double? EnglishMark
        {
            get
            {
                double j;
                if (double.TryParse(tbEnglishMark.Text.Trim(), out j))
                    return j;
                else
                    return null;
            }
            set { tbEnglishMark.Text = Util.ToStr(value); }
        }
        public bool StartEnglish
        {
            get { return chbStartEnglish.Checked; }
            set { chbStartEnglish.Checked = value; }
        }
        public bool EgeInSpbgu
        {
            get { return chbEgeInSpbgu.Checked; }
            set { chbEgeInSpbgu.Checked = value; }
        }
        public bool HasTRKI
        {
            get { return chbHasTRKI.Checked; }
            set { chbHasTRKI.Checked = value; }
        }
        public string TRKICertificateNumber
        {
            get { return tbTRKICertificateNumber.Text.Trim(); }
            set { tbTRKICertificateNumber.Text = value; }
        }
    }
}
