using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Priem
{
    public class ShortCompetition
    {
        public Guid Id { get; private set; }
        public Guid PersonId { get; private set; }
        public Guid EntryId { get; set; }
        public Guid CommitId { get; private set; }

        public int StudyLevelId { get; set; }
        public string StudyLevelName { get; set; }
        public int LicenseProgramId { get; set; }
        public string LicenseProgramName { get; set; }
        public int ObrazProgramId { get; set; }
        public string ObrazProgramName { get; set; }
        public Guid? ProfileId { get; set; }
        public string ProfileName { get; set; }
        public int FacultyId { get; set; }
        public string FacultyName { get; set; }

        public int StudyFormId { get; set; }
        public string StudyFormName { get; set; }
        public int StudyBasisId { get; set; }
        public string StudyBasisName { get; set; }

        public bool HasCompetition { get; set; }
        public int CompetitionId { get; set; }
        public string CompetitionName { get; set; }

        public ShortCompetition(Guid _Id, Guid _CommitId, Guid _EntryId, Guid _PersonId)
        {
            Id = _Id;
            CommitId = _CommitId;
            EntryId = _EntryId;
            PersonId = _PersonId;
        }
    }
}
