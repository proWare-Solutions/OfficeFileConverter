using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileConverter
{
  internal class ConvertException: Exception
  { 
    public ActionStatus Status { get;private set; }

    public ConvertException(string message, ActionStatus status):base(message) { 
      Status = status;
    }

    public ConvertException(string message, Exception innerException, ActionStatus status):base(message, innerException) {  
      Status = status; 
    }
  }
}
