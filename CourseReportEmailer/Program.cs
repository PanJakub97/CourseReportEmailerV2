using CourseReportEmailer.Models;
using CourseReportEmailer.Repository;
using CourseReportEmailer.Workers;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace CourseReportEmailer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {

                EnrollmentDetailReportCommand command = new EnrollmentDetailReportCommand(@"Data Source=WINDELL-EG5KDM3\SQLEXPRESS;Initial Catalog=CourseReport;Integrated Security=True");
                IList<EnrollmentDetailReportModel> models = command.GetList();

                var reportFileName = "EnrollmentDetailReport.xlsx";
                EnrollmentDetailReportSpeadSheetCreator enrollmentSheetCreator = new EnrollmentDetailReportSpeadSheetCreator();
                enrollmentSheetCreator.Create(reportFileName, models);

                EnrollmentDetailReportEmailSender emailer = new EnrollmentDetailReportEmailSender();
                emailer.Send(reportFileName);
            }
            catch (Exception ex)
            {

                Console.WriteLine("Something went wrong: {0}", ex.Message);
            }
        }
    }
}
  