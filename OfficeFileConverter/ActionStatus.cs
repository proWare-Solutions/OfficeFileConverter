using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileConverter
{
  internal enum ActionStatus
  {
    Unknown,
    Success,
    OriginalRemoveFailed,
    FileNotFound,
    AccessDenied,
    SaveFailed,
    OpenDatabaseFailed
  }
}
