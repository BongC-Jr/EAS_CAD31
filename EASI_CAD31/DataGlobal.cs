using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EASI_CAD31
{
   public class DataGlobal
   {
      public static SheetsService sheetsService { get; set; }

      public static string licSSheetsID { get; set; }

      public static string gpaSSheetID { get; set; }

      public static string slabGSheetViewer { get; set; }

      public static string credentialDirectory { get; set; }

      public static string credLicPath { get; set; }

      public static string bbsGSheet { get; set; }

      public static string bbsGSheetId { get; set; }

      public static string convofilepath { get; set; }

      public static string aiConvoGSheetId { get; set; }
      public static IList<IList<object>> trainingConversation { get; set; }

      public static IList<IList<object>> trainingConvCai2 { get; set; }

      public static bool isDevMessageOn { get; set; } = false;

      public static string UserMessage { get; set; }

    }
}
