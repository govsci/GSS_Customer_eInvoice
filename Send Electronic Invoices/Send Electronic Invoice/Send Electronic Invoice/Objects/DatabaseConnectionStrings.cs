using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Send_Electronic_Invoice.Objects
{
    public class DatabaseConnectionStrings
    {
        public static string PrdNavDb = "Data Source=PRD-NAV-DB;Initial Catalog=GSS;Integrated Security=SSPI;MultipleActiveResultSets=true; Asynchronous Processing=true;Connection Timeout=100000";
        public static string TstNavDb = "Data Source=TST-NAV-DB;Initial Catalog=GSS;User id=sjang;password=2878920;MultipleActiveResultSets=true; Asynchronous Processing=true;Connection Timeout=100000";
        public static string PrdEcomDb = "Data Source=PRD-ECOM-DB; user id=devuser; password=2878920; Initial Catalog=PRD-ECOM-DB; MultipleActiveResultSets=true; Asynchronous Processing=true;Connection Timeout=100000";
        public static string TstEcomDb = "Data Source=TST-ECOM-DB; user id=sjang; password=2878920; Initial Catalog=TST-ECOM-DB; MultipleActiveResultSets=true; Asynchronous Processing=true;Connection Timeout=100000";
    }
}
