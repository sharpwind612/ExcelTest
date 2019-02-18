using System.Data;
using System.IO;
using System.Text;
using Excel;
using UnityEditor;
using UnityEngine;

namespace EditorTool
{
    public class ExcelLoadTest : Editor
    {

        [MenuItem("Tools/打印表格")]
        public static void CreateItemAsset()
        {
            GameReadExcel("test.xlsx");
            Debug.Log("加载完毕!!!");
        }


        /// <summary>
        /// 只读Excel方法
        /// </summary>
        /// <param name="ExcelPath"></param>
        /// <returns></returns>
        public static void GameReadExcel(string ExcelPath)
        {
            FileStream stream = File.Open(Application.dataPath + "/Data/" + ExcelPath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            DataSet result = excelReader.AsDataSet();

            int columns = result.Tables[0].Columns.Count;//获取列数
            int rows = result.Tables[0].Rows.Count;//获取行数

            StringBuilder sb = new StringBuilder();
            //从第二行开始读
            for (int i = 1; i < rows; i++)
            {
                sb.Clear();
                for (int j = 0; j < columns; j++)
                {
                    sb.Append(result.Tables[0].Rows[i][j].ToString());
                    if (j != columns - 1)
                    {
                        sb.Append(",");
                    }
                }
                Debug.Log(sb.ToString());
            }

        }
    }
}