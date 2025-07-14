using System.IO;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using UglyToad.PdfPig;
using ClosedXML.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Collections.Concurrent;

namespace WordPdfSimilarCompare
{
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

                if (files.Count < 2)
                {
                    StatusText.Text = "目录中没有足够的 .docx 或 .pdf 文件进行比较";
                    return;
                }

                var texts = new ConcurrentDictionary<string, string>();

                // 并行提取文本内容
                await Task.Run(() =>
                {
                    Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, file =>
                    {
                        string text = file.EndsWith(".docx")
                            ? ExtractTextFromDocx(file)
                            : ExtractTextFromPdf(file);
                        texts[file] = text;

                        Dispatcher.Invoke(() =>
                        {
                            ProgressText.Text = $"读取：{Path.GetFileName(file)}";
                        });
                    });
                });

                var fileNames = texts.Keys.ToList();
                int totalComparisons = fileNames.Count * (fileNames.Count - 1) / 2;
                int comparisonCount = 0;

                var results = new ConcurrentBag<(string FileA, string FileB, double Similarity)>();

                // 并行比较相似度并更新进度条
                await Task.Run(() =>
                {
                    Parallel.For(0, fileNames.Count, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, i =>
                    {
                        for (int j = i + 1; j < fileNames.Count; j++)
                        {
                            var a = fileNames[i];
                            var b = fileNames[j];
                            double sim = CalculateLevenshteinSimilarity(texts[a], texts[b]);

                            results.Add((Path.GetFileName(a), Path.GetFileName(b), sim));

                            int count = Interlocked.Increment(ref comparisonCount);
                            double progress = (double)count / totalComparisons * 100;

                            Dispatcher.Invoke(() =>
                            {
                                ProgressBar.Value = progress;
                                ProgressText.Text = $"比较：{Path.GetFileName(a)} vs {Path.GetFileName(b)}";
                            });
                        }
                    });
                });

                string outputPath = Path.Combine(folder, "SimilarFile.xlsx");
                ExportToExcel(results.ToList(), outputPath);

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
            return string.Join("\n",
                doc.MainDocumentPart?.Document.Body?
                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Select(t => t.Text) ?? []);
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
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = s[i - 1] == t[j - 1] ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost
                    );
                }
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
            foreach (var r in results.OrderByDescending(r => r.Similarity))
            {
                ws.Cell(row, 1).Value = r.FileA;
                ws.Cell(row, 2).Value = r.FileB;
                ws.Cell(row, 3).Value = r.Similarity;
                row++;
            }

            ws.Columns().AdjustToContents();
            ws.Range("A1:C1").Style.Font.Bold = true;

            workbook.SaveAs(path);
        }
    }
}
