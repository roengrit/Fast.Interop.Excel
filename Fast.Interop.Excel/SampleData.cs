using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fast.Interop.Excel
{
    class SampleData
    {
        public DataTable GenData()
        {
            var Data = new DataTable();

            Data.Columns.Add("Running", typeof(int));
            Data.Columns.Add("TextVal", typeof(string));

            for (int i = 0; i < 1000000; i++)
            {
                var NewRow = Data.NewRow();
                NewRow["Running"] = i;
                NewRow["TextVal"] = Guid.NewGuid().ToString();
                Data.Rows.Add(NewRow);
            }
            return Data;
        }
    }
}
