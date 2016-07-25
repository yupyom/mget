using System;
using System.Diagnostics;
using OfficeOpenXml;  //EPPlusクラス
using OfficeOpenXml.Style;
//using OfficeOpenXml.Style.XmlAccess;
using System.IO;      //FileInfoクラスを使用
using System.Drawing; //参照設定に[System.Drawing]を追加する
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Sgml;


namespace WFA01
{
    public static class Converter
    {
        public static bool checkbox1_state;
        public static bool checkbox2_state;
        public static string file_path;
        public static Form1 f1 = (Form1)Application.OpenForms[0];

        public static void Convert(IProgress<int> progress, IProgress<string> logger)
        {
            Debug.WriteLine(checkbox1_state);
            Debug.WriteLine(checkbox2_state);
            Debug.WriteLine(file_path);

            string s = file_path;
            FileInfo input_file = new FileInfo(s);

            // outputファイルを作成（常に）
            string output_file_name = @"\output.xlsx";
            FileInfo output_file;
            if (!checkbox2_state)
            {
                output_file_name = @"\_output.xlsx";
            }
            var parent_path = System.IO.Directory.GetParent(file_path);
            Debug.WriteLine(parent_path);
            output_file = new FileInfo(parent_path + output_file_name);

            // ファイルのアクセス権限を見る
            FileAttributes fas = File.GetAttributes(output_file.ToString());
            bool bReadOnly = ((fas & FileAttributes.ReadOnly) == FileAttributes.ReadOnly);
            if (bReadOnly)
            {
                var desktop_path = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                output_file = new FileInfo(desktop_path + output_file_name);
            }
            if (output_file.Exists)
            {
                output_file.Delete();
                output_file = new FileInfo(parent_path + output_file_name);
            }

            Debug.WriteLine(output_file.FullName);
            Debug.WriteLine(output_file.Exists);

            // まず、入力ファイルを開く
            using (var package = new ExcelPackage(input_file))
            {
                if (checkbox2_state)
                {
                    // アウトプットファイルに書き出し
                    using (var package2 = new ExcelPackage(output_file))
                    {
                        // 入力シートを選択
                        var i_sheet = package.Workbook.Worksheets[1];
                        // 出力シートを作成
                        var o_sheet = package2.Workbook.Worksheets.Add("結果");

                        DoProcess(progress, logger, i_sheet, o_sheet);
                        package2.Save();

                    }
                }
                else
                {
                    // 入力シートを選択
                    var sheet = package.Workbook.Worksheets[1];
                    DoProcess(progress, logger, sheet, sheet);
                    package.SaveAs(output_file);
                }
            }
            // ファイルの置換
            if (!Converter.checkbox2_state)
            {
                input_file.Delete();
                output_file.MoveTo(file_path);
            }


        }

        // プロセスを実行
        private static void DoProcess(IProgress<int> progress, IProgress<string> logger, ExcelWorksheet i_sheet, ExcelWorksheet o_sheet)
        {
            var start_pos = 0;
            if (checkbox1_state)
            {
                o_sheet.Cells[1, 1].Value = "URL";
                o_sheet.Cells[1, 2].Value = "タイトル";
                o_sheet.Cells[1, 3].Value = "デスクリプション";
                o_sheet.Cells[1, 4].Value = "キーワード";
                var header_cells = o_sheet.Cells[1, 1, 1, 4];
                header_cells.Style.Font.Bold = true;
                start_pos = 1;
            }

            int num_of_row = i_sheet.Dimension.End.Row;
            foreach (var index in Enumerable.Range(1 + start_pos, num_of_row - start_pos))
            {
                if (!f1.isProgress)
                {
                    break;
                }

                //
                string url = i_sheet.GetValue(index, 1).ToString(); //これは保証できない
                string title = "";
                string description = "";
                string keywords = "";
                //
                logger.Report(index.ToString() + "/" + num_of_row.ToString() + ": " + url + " を解析中...");
                int percentage = index * 100 / num_of_row;
                progress.Report(percentage);

                //URLの確認
                Uri uri = null;
                HttpWebResponse res = null;
                if (Uri.TryCreate(url, UriKind.Absolute, out uri))
                {
                    //ページの存在の確認
                    var results = GetStatusAndContent(url, false);
                    if (300 <= results.StatusCode && results.StatusCode < 400)
                    {
                        title = "_リダイレクトされました_";
                        description = results.Url;
                        var error_cells = o_sheet.Cells[index, 1, index, 3];
                        error_cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        error_cells.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    }
                    else if (400 <= results.StatusCode && results.StatusCode < 500)
                    {
                        title = "_ページが見つかりません_";
                        description = "";
                        var error_cells = o_sheet.Cells[index, 1, index, 3];
                        error_cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        error_cells.Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }
                    else if (500 <= results.StatusCode)
                    {
                        title = "_サーバエラーです_";
                        description = "";
                        var error_cells = o_sheet.Cells[index, 1, index, 3];
                        error_cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        error_cells.Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }
                    else
                    {
                        //status 200

                        var xml = ParseHtmlFromText((string)results.Content);
                        if (xml == null) continue;

                        XNamespace ns = xml.Root.Name.Namespace;
                        foreach (var item in xml.Descendants(ns + "meta"))
                        {
                            XAttribute att = item.Attribute("name");
                            if (att != null)
                            {
                                if (att.Value == "description")
                                {
                                    description = item.Attribute("content").Value;
                                }
                                else if (att.Value == "keywords")
                                {
                                    keywords = item.Attribute("content").Value;
                                }
                            }
                            var elem = xml.Descendants(ns + "title");
                            var titleElem = elem.FirstOrDefault();
                            if (titleElem != null)
                            {
                                title = titleElem.Value;
                            }
                        }
                    }
                }
                else //不正なURI
                {
                    title = "_不正なURIです_";
                    description = "";
                    var error_cells = o_sheet.Cells[index, 1, index, 3];
                    error_cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    error_cells.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }

                o_sheet.Cells[index, 1].Value = url;
                o_sheet.Cells[index, 2].Value = title;
                o_sheet.Cells[index, 3].Value = description;
                o_sheet.Cells[index, 4].Value = keywords;
            }

            if (f1.processPercentage == 100)
            {
                logger.Report("終了しました！");
            }
            return;
        }

        private static dynamic GetStatusAndContent(string url, bool autoredirect)
        {

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.UserAgent = "Mozilla/5.0 (compatible; mgetbot/1.1)";
            req.AllowAutoRedirect = autoredirect;
            HttpWebResponse res = null; //init
            HttpStatusCode statusCode;
            string content = "";
            string finalUrl = url;

            try
            {
                res = (HttpWebResponse)req.GetResponse();
                statusCode = res.StatusCode;
                finalUrl = res.ResponseUri.ToString();
            }
            catch (WebException ex)
            {
                res = (HttpWebResponse)ex.Response;

                if (res != null)
                {
                    statusCode = res.StatusCode;
                    finalUrl = res.ResponseUri.ToString();
                }
                else
                {
                    throw; // サーバ接続不可などの場合は再スロー
                }
            }

            if (300 <= (int)statusCode && (int)statusCode < 400)
            {
                finalUrl = res.Headers["Location"];
                if (finalUrl.IndexOf("://", System.StringComparison.Ordinal) == -1)
                {
                    Uri u = new Uri(new Uri(url), finalUrl);
                    finalUrl = u.ToString();
                }
            }
            else if ((int)statusCode == 200)
            {
                Stream s = res.GetResponseStream();
                StreamReader sr = new StreamReader(s);
                Debug.WriteLine("Encode:" + sr.CurrentEncoding);
                content = sr.ReadToEnd();
                sr.Close();
            }

            if (res != null)
            {
                res.Close();
            }

            return new
            {
                Charset = res.CharacterSet,
                StatusCode = (int)statusCode,
                Url = finalUrl,
                Content = content
            };
        }


        private static XDocument ParseHtmlFromText(String text)
        {
            using (TextReader sr = new StringReader(text))
            {
                return ParseHtml(sr);
            }
        }

        private static XDocument ParseHtml(TextReader reader)
        {
            using (var sgmlReader = new SgmlReader { DocType = "HTML", CaseFolding = CaseFolding.ToLower })
            {
                try
                {
                    sgmlReader.InputStream = reader;
                    return XDocument.Load(sgmlReader);
                }
                catch (Exception e2)
                {
                    Debug.WriteLine(e2.Message);
                }
                return null;
            }
        }

        private static Encoding myDetectEncoding(byte[] data)
        {
            /*  バイト配列をASCIIエンコードで文字列に変換       */
            String s = Encoding.ASCII.GetString(data);
            /*  <meta>タグを抽出するための正規表現              */
            Match mymatch = Regex.Match(s,
                @"<meta[^>]*charset\s*=\s*""?([-_\w]+)""?",
                RegexOptions.IgnoreCase);
            String e = mymatch.Success ? mymatch.Groups[1].Value : "shift-jis";
            Debug.WriteLine(e);
            return Encoding.GetEncoding(e);
        }

    }
}