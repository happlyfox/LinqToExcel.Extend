using System;
using System.Collections.Generic;
using System.Linq;

namespace LinqToExcel.Validate
{
    /// <summary>
    /// 工作簿验证
    /// </summary>
    public class WorkBookValidate
    {
        public string FilePath { get; set; }

        private List<WorkSheetValidate> _workSheetList = new List<WorkSheetValidate>();

        public List<WorkSheetValidate> WorkSheetList
        {
            get { return _workSheetList; }
            set { _workSheetList = value; }
        }

        public WorkSheetValidate this[string sheetName]
        {
            get
            {
                foreach (WorkSheetValidate sheetParameterContainer in _workSheetList)
                {
                    if (sheetParameterContainer.SheetName.Equals(sheetName))
                    {
                        return sheetParameterContainer;
                    }
                }
                throw new Exception("工作表不存在");
            }
        }

        public WorkSheetValidate this[int sheetIndex]
        {
            get
            {
                foreach (WorkSheetValidate sheetParameterContainer in _workSheetList)
                {
                    if (sheetParameterContainer.SheetIndex.Equals(sheetIndex))
                    {
                        return sheetParameterContainer;
                    }
                }
                throw new Exception("工作表不存在");
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="filePath">路径</param>
        public WorkBookValidate(string filePath)
        {
            FilePath = filePath;
            var excel = new ExcelQueryFactory(filePath);
            List<string> worksheetNames = excel.GetWorksheetNames().ToList();

            int sheetIndex = 0;
            foreach (var sheetName in worksheetNames)
            {
                WorkSheetList.Add(new WorkSheetValidate(filePath, sheetName, sheetIndex++));
            }
        }
    }
}