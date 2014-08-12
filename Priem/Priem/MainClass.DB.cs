using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Objects;
using System.Data;

using BaseFormsLib;
using EducServLib;

namespace Priem
{
    public static partial class MainClass
    {
        public static bool GetIsOpen(string tableName, string itemId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                ObjectParameter entId = new ObjectParameter("result", typeof(bool));
                context.Get_IsOpen(tableName, itemId, entId);

                return Convert.ToBoolean(entId.Value);
            }
        }

        public static string GetIsOpenHolder(string tableName, string itemId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                ObjectParameter entId = new ObjectParameter("result", typeof(string));
                context.Get_OpenHolder(tableName, itemId, entId);

                return entId.Value.ToString();
            }
        }

        public static void SetIsOpen(string tableName, string itemId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                context.Set_IsOpen(tableName, itemId, true, userName);
            }
        }

        public static void DeleteIsOpen(string tableName, string itemId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                context.Set_IsOpen(tableName, itemId, false, null);
            }
        }

        public static void DeleteAllOpenByHolder()
        {
            using (PriemEntities context = new PriemEntities())
            {
                context.DeleteAllOpenByHolder(userName);
            }           
        }

        public static void Delete(string tableName, string itemId)
        {
            if (!IsReadOnly())
            {
                using (PriemEntities context = new PriemEntities())
                {
                    context.Delete(tableName, itemId);
                }                
            }
        }

        public static DateTime ToSmallDateTime(this DateTime dt)
        {
            return new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second);
        }

        //инициализация постороителя запросов для списка с фильтрами
        public static void InitQueryBuilder()
        {
            qBuilder = new QueryBuilder("ed.qAbiturient");

            qBuilder.AddQueryItem(new QueryItem("ed.extPerson", "ed.extPerson.PersonNum", "Ид_номер"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "ed.qAbiturient.RegNum", "Рег_номер"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.Surname", "Фамилия"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.Name", "Имя"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.SecondName", "Отчество"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.Surname + ' '+ed.Person.Name + ' ' + ed.Person.SecondName", "ФИО"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.BirthDate", "Дата_рождения"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.BirthPlace", "Место_рождения"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.PassportSeries", "Серия_паспорта"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.PassportNumber", "Номер_паспорта"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.PassportAuthor", "Кем_выдан_паспорт"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.PassportDate", "Дата_выдачи_паспорта"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", QueryBuilder.GetBoolField("ed.Person.Sex"), "Пол_мужской"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_Contacts", "ed.Person_Contacts.Phone", "Телефон"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_Contacts", "ed.Person_Contacts.Email", "Email"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_Contacts", "ed.Person_Contacts.Code+' '+ed.Person_Contacts.City+' '+ed.Person_Contacts.Street+(Case when ed.Person_Contacts.House = '' then '' else ' д.'+ed.Person_Contacts.House end)+(Case when ed.Person_Contacts.Korpus = '' then '' else ' к.'+ed.Person_Contacts.Korpus end)+(Case when ed.Person_Contacts.Flat = '' then '' else ' кв.'+ed.Person_Contacts.Flat end)", "Адрес_регистрации"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_Contacts", "ed.Person_Contacts.CodeReal+' '+ed.Person_Contacts.CityReal+' '+ed.Person_Contacts.StreetReal+(Case when ed.Person_Contacts.HouseReal = '' then '' else ' д.'+ed.Person_Contacts.HouseReal end)+(Case when ed.Person_Contacts.KorpusReal = '' then '' else ' к.'+ed.Person_Contacts.KorpusReal end)+(Case when ed.Person_Contacts.FlatReal = '' then '' else ' кв.'+ed.Person_Contacts.FlatReal end)", "Адрес_проживания"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", QueryBuilder.GetBoolField("ed.Person_AdditionalInfo.HostelAbit"), "Предоставлять_общежитие_поступление"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", QueryBuilder.GetBoolField("ed.Person_AdditionalInfo.HasAssignToHostel"), "Выдано_направление_на_поселение"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", QueryBuilder.GetBoolField("ed.Person_EducationInfo.IsExcellent"), "Медалист"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.SchoolCity", "Город_уч_заведения"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.SchoolNum", "Номер_школы"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.AttestatRegion", "Регион_аттестата"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.AttestatSeries", "Серия_атт"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.AttestatNum", "Номер_атт"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.SchoolAVG", "Средний_балл_атт"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.DiplomSeries", "Серия_диплома"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.DiplomNum", "Номер_диплома"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.PassportCode", "Код_подразделения_паспорта"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "ed.Person.PersonalCode", "Личный_код_паспорт"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_Contacts", "ed.Person_Contacts.Mobiles", "Доп_контакты"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "ed.Person_AdditionalInfo.PersonInfo", "Данные_о_родителях"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.StartEnglish", "Англ_с_нуля"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.EnglishMark", "Англ_оценка"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.HEQualification", "Квалификация"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.HighEducation", "Место_предыдущего_образования_маг"));

            if (MainClass.dbType == PriemType.PriemMag || MainClass.dbType == PriemType.PriemForeigners)
            {
                qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.HighEducation", "Название_уч_заведения"));
                qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.HEProfession", "Направление_подготовки"));
                qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.HEExitYear", "Год_выпуска"));
            }
            else
            {
                qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.SchoolName", "Название_уч_заведения"));
                qBuilder.AddQueryItem(new QueryItem("ed.Person_EducationInfo", "ed.Person_EducationInfo.SchoolExitYear", "Год_выпуска"));
            }


            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", QueryBuilder.GetBoolField("ed.Person_AdditionalInfo.HostelEduc"), "Предоставлять_общежитие_обучение"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.IsListener"), "Слушатель"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.WithHE"), "Имеющий_ВВ"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.IsSecond"), "Программы_для_ВО"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.IsReduced"), "Программы_сокр"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.IsParallel"), "Программы_парал"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.IsPaid"), "Оплатил"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.BackDoc"), "Забрал_док"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.HasOriginals"), "Подал_подлинники"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.Checked"), "Данные_проверены"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "ed.qAbiturient.StudyNumber", "Номер_зачетки"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "ed.qAbiturient.BackDocDate", "Дата_возврата_док"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "ed.qAbiturient.DocDate", "Дата_подачи_док"));
            //qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.AttDocOrigin"), "Поданы_подлинник_атт"));
            //qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", QueryBuilder.GetBoolField("ed.qAbiturient.EgeDocOrigin"), "Поданы_подлинники_ЕГЭ"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "ed.qAbiturient.Coefficient", "Коэффициент_полупрохода"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "ed.qAbiturient.Priority", "Приоритет"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "ed.qAbiturient.SessionAVG", "Средний_балл_сессии"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "ed.qAbiturient.StudentStatus", "Статус_студента"));

            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "(CASE WHEN ed.qAbiturient.IsGosLine = 0 THEN 'Нет' ELSE 'Да' END)", "Гослиния"));

            qBuilder.AddQueryItem(new QueryItem("ed.PassportType", "ed.PassportType.Name", "Тип_паспорта"));
            qBuilder.AddQueryItem(new QueryItem("ed.qPersonContactsCountry", "ISNULL(ed.qPersonContactsCountry.ForeignCountryName, ed.qPersonContactsCountry.CountryName)", "Страна"));
            qBuilder.AddQueryItem(new QueryItem("Nationality", "ISNULL(ForeignNationality.Name, Nationality.Name)", "Гражданство"));
            qBuilder.AddQueryItem(new QueryItem("ed.Region", "ed.Region.Name", "Регион"));
            qBuilder.AddQueryItem(new QueryItem("ed.Language", "ed.[Language].Name", "Ин_язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.SchoolType", "ed.SchoolType.Name", "Тип_уч_заведения"));
            qBuilder.AddQueryItem(new QueryItem("CountryEduc", "CountryEduc.Name", "Страна_получ_пред_образ"));
            qBuilder.AddQueryItem(new QueryItem("ed.extFBSStatus", "ed.extFBSStatus.FBSStatus", "Статус_ФБС"));
            qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", "(select MAX(ed.extOlympiads.OlympValueName) AS Value FROM ed.extOlympiads WHERE extOlympiads.AbiturientId = ed.qAbiturient.Id)", "Степень_диплома"));

            qBuilder.AddQueryItem(new QueryItem("HostelFaculty", "HostelFaculty.Name", "Факультет_выдавший_направление"));
            qBuilder.AddQueryItem(new QueryItem("ed.qFaculty", "ed.qFaculty.Name", "Факультет"));
            qBuilder.AddQueryItem(new QueryItem("ed.qLicenseProgram", "ed.qLicenseProgram.Name", "Направление"));
            qBuilder.AddQueryItem(new QueryItem("ed.qLicenseProgram", "ed.qLicenseProgram.Code", "Код_направления"));
            qBuilder.AddQueryItem(new QueryItem("ed.qObrazProgram", "ed.qObrazProgram.Name", "Образ_программа"));
            qBuilder.AddQueryItem(new QueryItem("ed.qProfile", "ed.qProfile.Name", "Профиль"));
            qBuilder.AddQueryItem(new QueryItem("ed.StudyForm", "ed.StudyForm.Name", "Форма_обучения"));
            qBuilder.AddQueryItem(new QueryItem("ed.StudyBasis", "ed.StudyBasis.Name", "Основа_обучения"));
            qBuilder.AddQueryItem(new QueryItem("ed.Competition", "ed.Competition.Name", "Тип_конкурса"));
            qBuilder.AddQueryItem(new QueryItem("OtherCompetition", "OtherCompetition.Name", "Доп_тип_конкурса"));
            qBuilder.AddQueryItem(new QueryItem("ed.CelCompetition", "ed.CelCompetition.Name", "Целевик_тип"));            

            // ЕГЭ
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Top(1) ed.EgeCertificate.Number as Number FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND Year=2010 ) ", "Свидетельство_ЕГЭ_2010"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Top(1) ed.EgeCertificate.Number as Number FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND Year=2011 )", "Свидетельство_ЕГЭ_2011"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Top(1) ed.EgeCertificate.Number as Number FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND Year=2012 )", "Свидетельство_ЕГЭ_2012"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", QueryBuilder.GetBoolField("(select Top(1) ed.EgeCertificate.IsImported FROM ed.EgeCertificate WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND Year=2012 )"), "Загружено_Свид-во_ЕГЭ_2012"));


            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=5)", "ЕГЭ_русск.язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=4)", "ЕГЭ_математика"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=2)", "ЕГЭ_физика"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=8)", "ЕГЭ_химия"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=3)", "ЕГЭ_биология"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=1)", "ЕГЭ_история"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=7)", "ЕГЭ_география"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=11)", "ЕГЭ_англ.яз"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=12)", "ЕГЭ_немец.язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=13)", "ЕГЭ_франц.язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=9)", "ЕГЭ_обществознание"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=6)", "ЕГЭ_литература"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=14)", "ЕГЭ_испан.язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select MAX (ed.EgeMark.value) as mark FROM ed.EgeMark INNER JOIN ed.EgeCertificate ON ed.EgeMark.EgeCertificateId = ed.EgeCertificate.Id WHERE ed.EgeCertificate.PersonId = ed.Person.Id AND EgeExamNameId=10)", "ЕГЭ_информатика"));

            //Олимпиады
            //qBuilder.AddQueryItem(new QueryItem("ed.Person", "case when(SELECT count(*) FROM Olympiads WHERE OlympLevelId=1 and Olympiads.AbiturientId=ed.qAbiturient.id)>0 then 'Да' else 'Нет' end", "Международная"));
            //qBuilder.AddQueryItem(new QueryItem("ed.Person", "case when(SELECT count(*) FROM Olympiads WHERE OlympLevelId=2 and Olympiads.AbiturientId=ed.qAbiturient.id)>0 then 'Да' else 'Нет' end", "Всероссийкая"));
            //qBuilder.AddQueryItem(new QueryItem("ed.Person", "case when(SELECT count(*) FROM Olympiads WHERE OlympLevelId=4 and Olympiads.AbiturientId=ed.qAbiturient.id)>0 then 'Да' else 'Нет' end", "Межвузовская"));
            //qBuilder.AddQueryItem(new QueryItem("ed.Person", "case when(SELECT count(*) FROM Olympiads WHERE OlympLevelId=5 and Olympiads.AbiturientId=ed.qAbiturient.id)>0 then 'Да' else 'Нет' end", "СПбГУ"));
            //qBuilder.AddQueryItem(new QueryItem("ed.Person", "case when(SELECT count(*) FROM Olympiads WHERE OlympLevelId=3 and Olympiads.AbiturientId=ed.qAbiturient.id)>0 then 'Да' else 'Нет' end", "Региональная"));
            //qBuilder.AddQueryItem(new QueryItem("ed.Person", "case when(SELECT count(*) FROM Olympiads WHERE OlympLevelId=9 and Olympiads.AbiturientId=ed.qAbiturient.id)>0 then 'Да' else 'Нет' end", "Школьников"));

            //Привилегии
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "case when (ed.Person_AdditionalInfo.Privileges & 1)>0 then 'Да' else 'Нет' end", "сирота"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "case when (ed.Person_AdditionalInfo.Privileges & 2)>0 then 'Да' else 'Нет' end", "чернобылец"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "case when (ed.Person_AdditionalInfo.Privileges & 4)>0 then 'Да' else 'Нет' end", "военнослужащий"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "case when (ed.Person_AdditionalInfo.Privileges & 16)>0 then 'Да' else 'Нет' end", "полусирота"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "case when (ed.Person_AdditionalInfo.Privileges & 32)>0 then 'Да' else 'Нет' end", "инвалид"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "case when (ed.Person_AdditionalInfo.Privileges & 64)>0 then 'Да' else 'Нет' end", "уч.боев."));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "case when (ed.Person_AdditionalInfo.Privileges & 128)>0 then 'Да' else 'Нет' end", "стажник"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "case when (ed.Person_AdditionalInfo.Privileges & 256)>0 then 'Да' else 'Нет' end", "реб.-сирота"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person_AdditionalInfo", "case when (ed.Person_AdditionalInfo.Privileges & 512)>0 then 'Да' else 'Нет' end", "огр.возможности"));
             
            //Протоколы
            qBuilder.AddQueryItem(new QueryItem("ed.extEnableProtocol", "ed.extEnableProtocol.Number", "Протокол_о_допуске"));
            qBuilder.AddQueryItem(new QueryItem("ed.extEntryView", "ed.extEntryView.Number", "Представление"));
            qBuilder.AddQueryItem(new QueryItem("ed.extEntryView", "ed.extEntryView.OrderNum", "Номер_приказа_о_зачислении"));
            qBuilder.AddQueryItem(new QueryItem("ed.extEntryView", "ed.extEntryView.OrderDate", "Дата_приказа_о_зачислении"));

            qBuilder.AddQueryItem(new QueryItem("ed.extEntryView", "ed.extEntryView.OrderNumFor", "Номер_приказа_о_зачислении_иностр"));
            qBuilder.AddQueryItem(new QueryItem("ed.extEntryView", "ed.extEntryView.OrderDateFor", "Дата_приказа_о_зачислении_иностр"));

            //Сумма баллов
            qBuilder.AddQueryItem(new QueryItem("ed.extAbitMarksSum", "ed.extAbitMarksSum.TotalSum", "Сумма_баллов"));


            //экзамены 
            DataSet dsExams = _bdc.GetDataSet("SELECT DISTINCT ed.extExamInEntry.ExamId AS Id, ed.extExamInEntry.ExamName AS Name FROM ed.extExamInEntry");

            foreach (DataRow dr in dsExams.Tables[0].Rows)
                qBuilder.AddQueryItem(new QueryItem("ed.qAbiturient", string.Format("(select Sum(qMark.Value) FROM ed.qMark INNER JOIN ed.extExamInEntry ON ed.qMark.ExamInEntryId =ed.extExamInEntry.Id WHERE AbiturientId = ed.qAbiturient.Id AND ed.extExamInEntry.ExamId={0})", dr["Id"]), dr["Name"].ToString()));
            
            // Оценки из аттестата
            //
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	1)", "Аттестат_алгебра"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	2)", "Аттестат_англ_язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	3)", "Аттестат_астрономия"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	4)", "Аттестат_биология"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	5)", "Аттестат_вселенная_чел"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	6)", "Аттестат_вс_история"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	7)", "Аттестат_география"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	8)", "Аттестат_геометрия"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	9)", "Аттестат_информатика"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	10)", "Аттестат_история_Спб"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	11)", "Аттестат_ист_России"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	12)", "Аттестат_литература"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	13)", "Аттестат_мировая_худ_культура"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	14)", "Аттестат_обществознание"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	15)", "Аттестат_ОБЖ"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	16)", "Аттестат_русск_язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	17)", "Аттестат_технология"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	18)", "Аттестат_физика"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	19)", "Аттестат_физ_культура"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	20)", "Аттестат_химия"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	21)", "Аттестат_экология"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	22)", "Аттестат_немецкий_язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	23)", "Аттестат_испанский_язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	24)", "Аттестат_французский_язык"));
            qBuilder.AddQueryItem(new QueryItem("ed.Person", "(select Min (ed.AttMarks.value) as mark FROM ed.AttMarks WHERE ed.AttMarks.PersonId = ed.Person.Id AND AttSubjectId=	25)", "Аттестат_итальянский_язык"));

            //JOIN-ы
            qBuilder.AddTableJoint("ed.extPerson", " LEFT JOIN ed.extPerson ON ed.qAbiturient.PersonId = ed.extPerson.Id ");
            qBuilder.AddTableJoint("ed.Person_EducationInfo", " LEFT JOIN ed.Person_EducationInfo ON ed.Person.Id = ed.Person_EducationInfo.PersonId ");
            qBuilder.AddTableJoint("ed.Person_Contacts", " LEFT JOIN ed.Person_Contacts ON ed.Person.Id = ed.Person_Contacts.PersonId ");
            qBuilder.AddTableJoint("ed.Person_AdditionalInfo", " LEFT JOIN ed.Person_AdditionalInfo ON ed.Person.Id = ed.Person_AdditionalInfo.PersonId ");
            qBuilder.AddTableJoint("ed.Person", " LEFT JOIN ed.Person ON ed.qAbiturient.PersonId = ed.Person.Id ");
            qBuilder.AddTableJoint("ed.PassportType", " LEFT JOIN ed.PassportType ON ed.PassportType.Id = ed.Person.PassportTypeId ");
            qBuilder.AddTableJoint("ed.qPersonContactsCountry", " LEFT JOIN ed.qPersonContactsCountry ON ed.Person.Id = ed.qPersonContactsCountry.PersonId ");
            qBuilder.AddTableJoint("Nationality", " LEFT JOIN ed.ForeignCountry AS ForeignNationality ON ed.Person.ForeignNationalityId = ForeignNationality.Id  LEFT JOIN ed.Country AS Nationality ON ed.Person.NationalityId = Nationality.Id ");
            qBuilder.AddTableJoint("ed.Region", " LEFT JOIN ed.Region ON ed.Person_Contacts.RegionId = ed.Region.Id ");
            qBuilder.AddTableJoint("ed.Language", " LEFT JOIN ed.[Language] ON ed.qAbiturient.LanguageId = ed.[Language].Id ");
            qBuilder.AddTableJoint("ed.SchoolType", " LEFT JOIN ed.SchoolType ON ed.extPerson.SchoolTypeId = ed.SchoolType.Id ");
            qBuilder.AddTableJoint("CountryEduc", " LEFT JOIN ed.Country AS CountryEduc ON ed.extPerson.CountryEducId = CountryEduc.Id ");
            qBuilder.AddTableJoint("HostelFaculty", " LEFT JOIN ed.SP_Faculty AS HostelFaculty ON ed.Person_AdditionalInfo.HostelFacultyId = HostelFaculty.Id ");
            qBuilder.AddTableJoint("ed.extFBSStatus", " LEFT JOIN ed.extFBSStatus ON ed.extFBSStatus.PersonId = ed.Person.Id ");

            qBuilder.AddTableJoint("ed.qFaculty", " LEFT JOIN ed.qFaculty ON ed.qFaculty.Id = ed.qAbiturient.FacultyId ");
            qBuilder.AddTableJoint("ed.qLicenseProgram", " LEFT JOIN ed.qLicenseProgram ON ed.qLicenseProgram.Id = ed.qAbiturient.LicenseProgramId ");
            qBuilder.AddTableJoint("ed.qObrazProgram", " LEFT JOIN ed.qObrazProgram ON ed.qObrazProgram.Id = ed.qAbiturient.ObrazProgramId ");
            qBuilder.AddTableJoint("ed.qProfile", " LEFT JOIN ed.qProfile ON ed.qProfile.Id = ed.qAbiturient.ProfileId ");
            qBuilder.AddTableJoint("ed.StudyBasis", " LEFT JOIN ed.StudyBasis ON ed.StudyBasis.Id = ed.qAbiturient.StudyBasisId ");
            qBuilder.AddTableJoint("ed.StudyForm", " LEFT JOIN ed.StudyForm ON ed.StudyForm.Id = ed.qAbiturient.StudyFormId ");
            qBuilder.AddTableJoint("ed.Competition", " LEFT JOIN ed.Competition ON ed.Competition.Id = ed.qAbiturient.CompetitionId ");
            qBuilder.AddTableJoint("OtherCompetition", " LEFT JOIN ed.Competition AS OtherCompetition ON ed.qAbiturient.OtherCompetitionId = OtherCompetition.Id ");
            qBuilder.AddTableJoint("ed.CelCompetition", " LEFT JOIN ed.CelCompetition ON ed.qAbiturient.CelCompetitionId = ed.CelCompetition.Id ");
           
            qBuilder.AddTableJoint("ed.extEnableProtocol", " LEFT JOIN ed.extEnableProtocol ON ed.extEnableProtocol.AbiturientId = ed.qAbiturient.Id ");
            qBuilder.AddTableJoint("ed.extEntryView", " LEFT JOIN ed.extEntryView ON ed.extEntryView.AbiturientId = ed.qAbiturient.Id ");            
            qBuilder.AddTableJoint("ed.extAbitMarksSum", " LEFT JOIN ed.extAbitMarksSum ON ed.extAbitMarksSum.Id = ed.qAbiturient.Id ");
        }
    }
}
