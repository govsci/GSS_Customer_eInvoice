using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Send_Electronic_Invoice.Objects
{
    public class CodeError
    {
        public CodeError(Exception ex, string clss, string function, SqlCommand cmd) 
        {
            Error = ex;
            Class = clss;
            Function = function;
            CMD = cmd;
        }

        public Exception Error { get; }
        public string Class { get; }
        public string Function { get; }
        public SqlCommand CMD { get; }
    }
}
