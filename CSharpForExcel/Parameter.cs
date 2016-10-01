using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharpForExcel
{
    /// <summary>
    /// スクリプト実行時に渡す実行側との通信用クラス
    /// このオブジェクトを実行時に渡すとスクリプト内でReadCellとWriteCellメソッドを呼び出せる
    /// </summary>
    public class Parameter
    {
        /// <summary>
        /// 指定したセルをstring型で読み込む
        /// </summary>
        /// <param name="cellName">読み込むセルの名前("B2"のような指定方法)</param>
        /// <returns></returns>
        public string ReadCell(string cellName)
        {
            try
            {
                Excel.Worksheet activeSheet = ((Excel.Worksheet)ThisAddIn.application.ActiveSheet);
                var range = activeSheet.Range[cellName];
                object cell = activeSheet.Cells[range.Row, range.Column].Value;
                return cell.ToString();
            }catch(Exception e)
            {
                MessageBox.Show("エラーが発生しました。ReadCellメソッドのセル名が不正な可能性があります");
            }

            return string.Empty;
        }

        /// <summary>
        /// 指定したセルに値を書き込む
        /// </summary>
        /// <param name="cellName">書き込むセルの名前("B2"のような名前)</param>
        /// <param name="value">書き込むセルの内容</param>
        public void WriteCell(string cellName, object value)
        {
            try {
                Excel.Worksheet activeSheet = ((Excel.Worksheet)ThisAddIn.application.ActiveSheet);
                var range = activeSheet.Range[cellName];
                activeSheet.Cells[range.Row, range.Column] = value.ToString();
            }
            catch (Exception e)
            {
                MessageBox.Show("エラーが発生しました。WriteCellメソッドのセル名が不正な可能性があります");
            }
        }
    }
}
