using Excel = Microsoft.Office.Interop.Excel;

namespace СhangePassword
{
    public static class NewPassword
    {
        public static void Run()
        {
            const string PathOut = @"C:\Users\aleskovska\Desktop\Данные\ГенераторПароля.txt";
            const string PathOutHistory = @"C:\Users\aleskovska\Desktop\Данные\ГенераторПароляИстория.txt";

            string password = GetOldPassword(PathOut);
            string newPassword = GetNewPassword(PathOut, PathOutHistory);

            var directoryStart = new DirectoryInfo(@"\\192.168.0.200\share\Users\Аналитика");
            Recursion(directoryStart, password, newPassword);
        }

        private static void Base(DirectoryInfo directory, string password, string newPassword)
        {
            var files = directory.GetFiles("*.xls*");

            if (files.Count() > 0)
            {
                foreach (var file in files)
                {
                    bool flag;
                    ChangePassword(file.FullName, password, newPassword, out flag);

                    if (flag == false)
                        Console.WriteLine($"Ошибка: пароль не изменен {file.FullName}");
                }
            }
        }

        private static void Recursion(DirectoryInfo directory, string password, string newPassword)
        {
            var directories = directory.GetDirectories();
            Base(directory, password, newPassword);

            if (directories.Count() > 0)
            {
                foreach (var dir in directories)
                    Recursion(dir, password, newPassword);
            }
        }

        private static void ChangePassword(string pathFile, string password, string newPassword, out bool flag)
        {
            Excel.Application app = new Excel.Application();
            app.DisplayAlerts = false;
            flag = false;

            try
            {
                Excel.Workbook workBook = app.Workbooks.Open(pathFile, 0, false, 5, password, "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                workBook.Password = newPassword;

                workBook.SaveAs(pathFile, Type.Missing, newPassword, Type.Missing, Type.Missing, Type.Missing,
                                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                workBook.Close(false, Type.Missing, Type.Missing);
                workBook = null;
                flag = true;
            }
            catch
            {
                app.Workbooks.Close();
            }

            app.Quit();
            app = null;
            GC.Collect();
        }

        private static string GeneratePass()
        {
            string iPass = "";
            string[] arr = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "B", "C", "D", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "V", "W", "X", "Z", "b", "c", "d", "f", "g", "h", "j", "k", "m", "n", "p", "q", "r", "s", "t", "v", "w", "x", "z", "A", "E", "U", "Y", "a", "e", "i", "o", "u", "y" };
            Random rnd = new Random();
            for (int i = 0; i < 30; i = i + 1)
            {
                iPass = iPass + arr[rnd.Next(0, 57)];
            }

            return iPass;
        }

        private static string GetOldPassword( string PathOut)
        {
            using (var reader = new StreamReader(PathOut))
                return reader.ReadToEnd();
        }

        private static string GetNewPassword(string PathOut, string PathOutHistory)
        {
            string newPassword = GeneratePass();

            using (var writer = new StreamWriter(PathOut))
                writer.Write(newPassword);

            File.AppendAllText(PathOutHistory, "\n" + newPassword);
            return newPassword;
        }
    }
}
