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

namespace Group4337
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
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
                                Role = sheet.Cells[row, 1].Text,
                                FullName = sheet.Cells[row, 2].Text,
                                Login = sheet.Cells[row, 3].Text,
                                Password = sheet.Cells[row, 4].Text,
                            };
                            if (!string.IsNullOrWhiteSpace(emp.Login))
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

                var groups = employees.GroupBy(x => x.Role);
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
                            sheet.Cells[row, 1].Value = emp.Login;
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
    }
}



   