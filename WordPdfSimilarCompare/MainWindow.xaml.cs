using System.IO;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using UglyToad.PdfPig;
using ClosedXML.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace WordPdfSimilarCompare
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog { IsFolderPicker = true };
            if (dlg.ShowDialog() != CommonFileDialogResult.Ok) return;

            string folder = dlg.FileName!;
            StatusText.Text = $"正在分析目录: {folder} ...";

            ProgressText.Text = "";
            ProgressBar.Value = 0;

            try
            {
                var files = Directory.GetFiles(folder)
                    .Where(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                var texts = new Dictionary<string, string>();

                int fileIndex = 0;
                foreach (var file in files)
                {
                    fileIndex++;
                    ProgressText.Text = $"正在读取文件 {fileIndex}/{files.Count}: {Path.GetFileName(file)}";
                    await Task.Delay(50); // 模拟延迟，防止 UI 卡顿

                    if (file.EndsWith(".docx"))
                        texts[file] = ExtractTextFromDocx(file);
                    else if (file.EndsWith(".pdf"))
                        texts[file] = ExtractTextFromPdf(file);
                }

                var fileNames = texts.Keys.ToList();
                int totalComparisons = fileNames.Count * (fileNames.Count - 1) / 2;
                int comparisonCount = 0;

                var results = new List<(string FileA, string FileB, double Similarity)>();

                for (int i = 0; i < fileNames.Count; i++)
                {
                    for (int j = i + 1; j < fileNames.Count; j++)
                    {
                        var a = fileNames[i];
                        var b = fileNames[j];
                        double sim = CalculateLevenshteinSimilarity(texts[a], texts[b]);

                        results.Add((Path.GetFileName(a), Path.GetFileName(b), sim));

                        comparisonCount++;
                        double progress = (double)comparisonCount / totalComparisons * 100;
                        ProgressBar.Value = progress;
                        ProgressText.Text = $"正在比较：{Path.GetFileName(a)} vs {Path.GetFileName(b)}";
                        await Task.Delay(30); // 模拟小延迟以更新 UI
                    }
                }

                string outputPath = Path.Combine(folder, "SimilarFile.xlsx");
                ExportToExcel(results, outputPath);

                StatusText.Text = $"分析完成，结果已保存到：{outputPath}";
                ProgressText.Text = "完成";
                ProgressBar.Value = 100;

            }
            catch (Exception ex)
            {
                StatusText.Text = $"出错：{ex.Message}";
            }
        }

        string ExtractTextFromDocx(string path)
        {
            using var doc = WordprocessingDocument.Open(path, false);
            return doc.MainDocumentPart?.Document.Body?.InnerText ?? "";
        }

        string ExtractTextFromPdf(string path)
        {
            using var pdf = PdfDocument.Open(path);
            return string.Join("\n", pdf.GetPages().Select(p => p.Text));
        }

        double CalculateLevenshteinSimilarity(string s1, string s2)
        {
            int distance = LevenshteinDistance(s1, s2);
            int maxLength = Math.Max(s1.Length, s2.Length);
            if (maxLength == 0) return 1.0;
            return 1.0 - (double)distance / maxLength;
        }

        int LevenshteinDistance(string s, string t)
        {
            int n = s.Length, m = t.Length;
            int[,] d = new int[n + 1, m + 1];
            for (int i = 0; i <= n; i++) d[i, 0] = i;
            for (int j = 0; j <= m; j++) d[0, j] = j;

            for (int i = 1; i <= n; i++)
                for (int j = 1; j <= m; j++)
                {
                    int cost = s[i - 1] == t[j - 1] ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }

            return d[n, m];
        }

        void ExportToExcel(List<(string FileA, string FileB, double Similarity)> results, string path)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Similarity");

            ws.Cell(1, 1).Value = "文件A";
            ws.Cell(1, 2).Value = "文件B";
            ws.Cell(1, 3).Value = "相似度 (0~1)";

            int row = 2;
            foreach (var r in results)
            {
                ws.Cell(row, 1).Value = r.FileA;
                ws.Cell(row, 2).Value = r.FileB;
                ws.Cell(row, 3).Value = r.Similarity;
                row++;
            }

            workbook.SaveAs(path);
        }
    }
}