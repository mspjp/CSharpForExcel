using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using System.Reflection;
using System.IO;

namespace CSharpForExcel
{
    public partial class MainRibbon
    {
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// runボタンがクリックされたときのイベントハンドラ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void buttonRun_Click(object sender, RibbonControlEventArgs e)
        {
            //選択されているセルを取得する
            Excel.Range rng = (Excel.Range)ThisAddIn.application.Selection;
            var count = int.Parse(rng.Count.ToString());
            if (count == 0)
            {
                MessageBox.Show("C#が書かれたセルを選択してください", "エラー");
                return;
            }

            //選択されたセルをスクリプトコードに直す
            string code = string.Empty;
            if (count == 1)
            {
                code = rng.Value;
            }
            else
            {
                var list = new List<string>();
                foreach (var cell in rng.Value)
                {
                    object cellObj = cell;
                    if (cellObj != null)
                    {
                        list.Add(cellObj.ToString());
                    }
                }

                code = string.Join("\n", list);

            }

            //スクリプトにSystem名前空間をusingしている状態で実行できるように設定
            ScriptOptions options = ScriptOptions.Default.AddImports("System");
            //実行ホストとの通信用オブジェクト
            var param = new Parameter();

            //スクリプトを作成
            var script = CSharpScript.Create(code, globalsType: typeof(Parameter),options:options);
            
            try
            {
                //Roslynを利用してスクリプトを実行する
                var state = await script.RunAsync(param);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "コンパイルに失敗しました");
            }
            
        }
    }
}
