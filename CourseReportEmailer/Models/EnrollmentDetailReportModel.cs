using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CourseReportEmailer.Models
{
    internal class EnrollmentDetailReportModel
    {
		public int EnrollmentId { get; set; }

        public string FirstName { get; set; }

		public string LastName { get; set; }

        public string CourseCode { get; set; }

        public string Description { get; set; }

        public string Degre { get; set; }

		public string Pseudo { get; set; }

	}
}
