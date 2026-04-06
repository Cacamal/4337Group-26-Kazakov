using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using OfficeOpenXml;
using System.Security.Cryptography;
using System.Reflection.Metadata;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;


namespace Group4337
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadDataFromDatabase();
        }
        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            var openFile = new OpenFileDialog { Filter = "Excel|*.xlsx" };
            if (openFile.ShowDialog() == true)
            {
                using (var db = new ApplicationDbContext())
                {
                    db.Database.EnsureCreated();
                    db.Employees.RemoveRange(db.Employees);

                    using (var package = new ExcelPackage(new FileInfo(openFile.FileName)))
                    {
                        var sheet = package.Workbook.Worksheets[0];
                        for (int row = 2; row <= sheet.Dimension.Rows; row++)
                        {
                            var emp = new Employee
                            {
                                Position = sheet.Cells[row, 1].Text,
                                FullName = sheet.Cells[row, 2].Text,
                                Log = sheet.Cells[row, 3].Text,
                                Password = sheet.Cells[row, 4].Text,
                            };
                            if (!string.IsNullOrWhiteSpace(emp.Log))
                                db.Employees.Add(emp);
                        }
                        db.SaveChanges();
                    }
                    dgEmployees.ItemsSource = db.Employees.ToList();
                    txtStatus.Text = $"Загружено {db.Employees.Count()} записей";
                }
            }
        }
        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new ApplicationDbContext())
            {
                var employees = db.Employees.ToList();
                if (!employees.Any()) return;

                var groups = employees.GroupBy(x => x.Position);
                using (var package = new ExcelPackage())
                {
                    foreach (var group in groups)
                    {
                        string sheetName = group.Key.Length > 31 ? group.Key.Substring(0, 31) : group.Key;
                        var sheet = package.Workbook.Worksheets.Add(sheetName);
                        sheet.Cells[1, 1].Value = "Логин";
                        sheet.Cells[1, 2].Value = "Пароль";
                        int row = 2;
                        foreach (var emp in group)
                        {
                            sheet.Cells[row, 1].Value = emp.Log;
                            sheet.Cells[row, 2].Value = HashPassword(emp.Password);
                            row++;
                        }
                        sheet.Cells.AutoFitColumns();
                    }
                    var saveFile = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = "Employees_by_Role.xlsx" };
                    if (saveFile.ShowDialog() == true)
                        package.SaveAs(new FileInfo(saveFile.FileName));
                }
            }
        }
        public static string HashPassword(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(password);
                byte[] hashBytes = sha256.ComputeHash(bytes);

                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                    builder.Append(hashBytes[i].ToString("x2"));

                return builder.ToString();
            }
        }






        private void LoadDataFromDatabase()
        {
            using (var db = new ApplicationDbContext())
            {
                db.Database.EnsureCreated();
                var employees = db.Employees.ToList();
                dgEmployees.ItemsSource = employees;
                txtStatus.Text = $"Загружено {employees.Count()} записей из БД";
            }
        }
        // Импорт JSON данных (вариант 10 - файл 5.json)
        private void btnImportJson_Click(object sender, RoutedEventArgs e)
        {
            var openFile = new OpenFileDialog
            {
                Filter = "JSON файлы (*.json)|*.json|Все файлы (*.*)|*.*",
                Title = "Выберите файл 5.json"
            };

            if (openFile.ShowDialog() == true)
            {
                try
                {
                    string jsonContent = File.ReadAllText(openFile.FileName);

                    // Десериализация прямо в модель Employee (поля совпадают!)
                    var employees = JsonSerializer.Deserialize<List<Employee>>(jsonContent, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    });

                    if (employees == null || !employees.Any())
                    {
                        MessageBox.Show("JSON файл не содержит данных или имеет неверный формат");
                        return;
                    }

                    using (var db = new ApplicationDbContext())
                    {
                        db.Database.EnsureCreated();

                        // Очищаем старые данные
                        db.Employees.RemoveRange(db.Employees);

                        // Хэшируем пароли перед сохранением
                        foreach (var emp in employees)
                        {
                            emp.Password = HashPassword(emp.Password);
                            db.Employees.Add(emp);
                        }

                        db.SaveChanges();
                    }

                    // Обновляем DataGrid
                    LoadDataFromDatabase();
                    txtStatus.Text = $"Импортировано {employees.Count()} сотрудников из JSON";
                    MessageBox.Show($"Успешно импортировано {employees.Count()} записей!", "Успех",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (JsonException ex)
                {
                    MessageBox.Show($"Ошибка парсинга JSON: {ex.Message}\n\n" +
                                    $"Убедитесь, что файл имеет правильный формат:\n" +
                                    $"[{{\"Id\":1,\"Position\":\"Администратор\",\"FullName\":\"ФИО\",\"Log\":\"email\",\"Password\":\"pass\"}}]",
                                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка импорта JSON: {ex.Message}", "Ошибка",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void btnExportWord_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new ApplicationDbContext())
            {
                var employees = db.Employees.ToList();

                if (!employees.Any())
                {
                    MessageBox.Show("Нет данных для экспорта. Сначала импортируйте JSON файл.",
                                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Группировка по должности (Position) - это критерий из варианта 10
                var groups = employees.GroupBy(x => x.Position).OrderBy(g => g.Key);

                using (var memoryStream = new MemoryStream())
                {
                    using (var wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
                    {
                        var mainPart = wordDocument.AddMainDocumentPart();
                        mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                        var body = mainPart.Document.AppendChild(new Body());

                        // Заголовок документа
                        var titleParagraph = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                        var titleRun = titleParagraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
                        titleRun.AppendChild(new Text("Сотрудники по должностям"));
                        titleRun.RunProperties = new RunProperties { Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold() };
                        titleParagraph.ParagraphProperties = new ParagraphProperties
                        {
                            Justification = new Justification() { Val = JustificationValues.Center },
                            SpacingBetweenLines = new SpacingBetweenLines() { After = "240" }
                        };

                        int groupCount = 0;
                        foreach (var group in groups)
                        {
                            groupCount++;

                            // Заголовок категории (должность)
                            var categoryParagraph = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                            var categoryRun = categoryParagraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
                            categoryRun.AppendChild(new Text($"Должность: {group.Key}"));
                            categoryRun.RunProperties = new RunProperties { Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold() };
                            categoryParagraph.ParagraphProperties = new ParagraphProperties
                            {
                                SpacingBetweenLines = new SpacingBetweenLines() { Before = "240", After = "120" }
                            };

                            // Создание таблицы
                            var table = new DocumentFormat.OpenXml.Wordprocessing.Table();

                            // Стиль таблицы
                            var tableProperties = new TableProperties(
                                new TableBorders(
                                    new TopBorder() { Val = BorderValues.Single, Size = 1 },
                                    new BottomBorder() { Val = BorderValues.Single, Size = 1 },
                                    new LeftBorder() { Val = BorderValues.Single, Size = 1 },
                                    new RightBorder() { Val = BorderValues.Single, Size = 1 },
                                    new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 1 },
                                    new InsideVerticalBorder() { Val = BorderValues.Single, Size = 1 }
                                ),
                                new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Dxa }
                            );
                            table.AppendChild(tableProperties);

                            // Заголовки таблицы (Логин, Пароль - согласно варианту 10)
                            var headerRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                            headerRow.AppendChild(CreateTableCell("Логин (Email)", true));
                            headerRow.AppendChild(CreateTableCell("Пароль (хэш)", true));
                            table.AppendChild(headerRow);

                            // Данные сотрудников в группе
                            foreach (var emp in group)
                            {
                                var dataRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                                dataRow.AppendChild(CreateTableCell(emp.Log, false));
                                dataRow.AppendChild(CreateTableCell(emp.Password, false));
                                table.AppendChild(dataRow);
                            }

                            body.AppendChild(table);

                            // Добавляем информацию о количестве сотрудников
                            var countParagraph = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                            var countRun = countParagraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
                            countRun.AppendChild(new Text($"Всего сотрудников в должности '{group.Key}': {group.Count()}"));
                            countRun.RunProperties = new RunProperties { Italic = new DocumentFormat.OpenXml.Wordprocessing.Italic() };
                            countParagraph.ParagraphProperties = new ParagraphProperties
                            {
                                SpacingBetweenLines = new SpacingBetweenLines() { Before = "120", After = "120" }
                            };

                            // Разрыв страницы (кроме последней группы)
                            if (groupCount < groups.Count())
                            {
                                var breakParagraph = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                                var breakRun = breakParagraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
                                breakRun.AppendChild(new Break() { Type = BreakValues.Page });
                            }
                        }
                    }

                    // Сохранение файла
                    var saveFile = new SaveFileDialog
                    {
                        Filter = "Word Document (*.docx)|*.docx",
                        FileName = "Employees_by_Position.docx"
                    };

                    if (saveFile.ShowDialog() == true)
                    {
                        File.WriteAllBytes(saveFile.FileName, memoryStream.ToArray());
                        txtStatus.Text = $"Экспорт в Word завершен. Файл: {saveFile.FileName}";
                        MessageBox.Show($"Документ успешно сохранен:\n{saveFile.FileName}", "Успех",
                                        MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
        }
        private DocumentFormat.OpenXml.Wordprocessing.TableCell CreateTableCell(string text, bool isHeader = false)
        {
            var cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
            var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            var run = new DocumentFormat.OpenXml.Wordprocessing.Run();
            run.AppendChild(new Text(text));

            if (isHeader)
            {
                run.RunProperties = new RunProperties { Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold() };
            }

            paragraph.AppendChild(run);
            cell.AppendChild(paragraph);

            cell.TableCellProperties = new TableCellProperties(
                new TableCellMargin()
                {
                    TopMargin = new TopMargin() { Width = "50", Type = TableWidthUnitValues.Dxa },
                    BottomMargin = new BottomMargin() { Width = "50", Type = TableWidthUnitValues.Dxa }
                }
            );

            return cell;
        }
    }
}



   