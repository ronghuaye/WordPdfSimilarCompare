# Word/PDF 文本相似度比较工具 (Word/PDF Similarity Comparer)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

一个简单而强大的WPF桌面应用程序，用于分析指定文件夹内所有 `.docx` 和 `.pdf` 文件的文本内容，计算它们两两之间的相似度，并将结果导出为一份清晰的 Excel 报告。


<img width="386" height="233" alt="image" src="https://github.com/user-attachments/assets/f392ff0d-576f-478d-b644-7eb449b8fd89" />


---

## ✨ 功能特性 (Features)

*   **支持多种格式**: 同时支持微软 Word (`.docx`) 和 PDF (`.pdf`) 文件。
*   **批量比较**: 自动发现并比较指定文件夹内的所有支持文件。
*   **高效并行处理**: 利用多核处理器并行读取文件和计算相似度，大大缩短了处理时间。
*   **精准相似度算法**: 基于经典的 **Levenshtein 距离（编辑距离）** 算法来量化两个文档内容的相似程度。
*   **实时进度反馈**: 通过进度条和状态文本实时显示分析进度，用户体验良好。
*   **清晰的Excel报告**: 将所有比较结果（文件A, 文件B, 相似度）按相似度从高到低排序，并导出为 `.xlsx` 文件，方便筛选和分析。
*   **响应式UI**: 整个分析过程在后台线程中运行，确保用户界面在处理大量文件时也不会卡顿。

## 🚀 如何使用 (Getting Started)

1.  **准备文件**: 将所有需要比较的 `.docx` 和 `.pdf` 文件放在同一个文件夹中。
2.  **运行程序**: 打开本应用程序 (`WordPdfSimilarCompare.exe`)。
3.  **选择目录**: 点击 "选择目录并开始分析" 按钮，然后在弹出的对话框中选择您刚刚准备好的文件夹。
4.  **等待分析**: 程序将自动开始分析。您可以在界面上看到实时进度，例如正在读取哪个文件，或正在比较哪两个文件。
5.  **查看结果**: 分析完成后，程序会提示 "分析完成"。此时，您可以到您选择的那个文件夹中，找到一个名为 `SimilarFile.xlsx` 的 Excel 文件，其中包含了所有的比较结果。

## 🔧 技术栈 (Technologies Used)

*   **框架**: .NET / WPF (Windows Presentation Foundation)
*   **Word 文档读取**: [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) (`DocumentFormat.OpenXml.Packaging`)
*   **PDF 文档读取**: [PdfPig](https://github.com/UglyToad/PdfPig)
*   **Excel 文件生成**: [ClosedXML](https://github.com/ClosedXML/ClosedXML)
*   **现代文件夹对话框**: [Windows API Code Pack](https://github.com/aybe/Windows-API-Code-Pack-1.1)
*   **并发处理**: TPL (Task Parallel Library) - `Task.Run`, `Parallel.ForEach`, `ConcurrentDictionary`, `Interlocked`

## ⚙️ 工作原理 (How It Works)

本工具的工作流程分为以下几个关键步骤：

1.  **文件发现**: 用户选择一个目录后，程序会递归扫描该目录，找出所有以 `.docx` 或 `.pdf` 结尾的文件。
2.  **并行文本提取**:
    *   程序利用 `Parallel.ForEach` 对所有发现的文件进行并行处理。
    *   对于 `.docx` 文件，使用 `Open XML SDK` 提取其主体部分的所有文本。
    *   对于 `.pdf` 文件，使用 `PdfPig` 库逐页提取文本。
    *   所有提取出的文本内容被存储在一个线程安全的 `ConcurrentDictionary` 中，以文件名作为键，文本内容作为值。
3.  **并行相似度计算**:
    *   程序会生成一个包含所有文件两两组合的比较任务列表。
    *   使用嵌套的 `Parallel.For` 循环来并行处理这些比较任务。
    *   对于每一对文件，从内存中获取它们的文本内容，并调用 `CalculateLevenshteinSimilarity` 方法计算相似度。该方法基于编辑距离，计算公式为 `1 - (编辑距离 / 两个文本中的最大长度)`，结果在 `0` (完全不同) 到 `1` (完全相同) 之间。
    *   每次比较的进度都通过 `Interlocked.Increment` 进行线程安全的更新，并实时反馈到UI的进度条上。
4.  **结果导出**:
    *   所有的比较结果（文件A、文件B、相似度得分）被收集到一个线程安全的 `ConcurrentBag` 中。
    *   分析全部结束后，程序使用 `ClosedXML` 库将收集到的结果按相似度降序排列，并写入一个格式化的 Excel 文件中。

## 🤝 贡献 (Contributing)

欢迎提交 Pull Request 或创建 Issue 来帮助改进这个项目。

1.  Fork 本项目
2.  创建您的功能分支 (`git checkout -b feature/AmazingFeature`)
3.  提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4.  推送到分支 (`git push origin feature/AmazingFeature`)
5.  打开一个 Pull Request

## 📄 许可证 (License)

本项目采用 MIT 许可证。详情请见 [LICENSE](LICENSE) 文件。
