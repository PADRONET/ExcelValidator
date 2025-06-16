using System.Text.RegularExpressions;

namespace ExcelValidator.Functions
{
    public static class AsinValidator
    {
        public static bool IsValidAsin(this string asin)
        {
            return new Regex("^B\\d{2}\\w{7}|\\d{9}(X|\\d)$").IsMatch(asin);
        }
    }

    public static class EanValidator
    {
        public static bool IsValidEan(this string ean)
        {
            if (ean.IsValidSemGetIn())
                return true;
            if (ean.IsValidEanFormat())
            {
                return ean.IsValidEanChecksum();
            }

            return false;
        }

        public static bool IsValidEanFormat(this string ean)
        {
            return ean.Length == 13;
        }

        public static bool IsValidEanChecksum(this string ean)
        {
            try
            {
                return ValidatorHelpers.IsStandardChecksumValid(ean);
            }
            catch
            {
                return false;
            }
        }

        public static bool IsValidSemGetIn(this string ean)
        {
            try
            {
                if (ean.ToUpper().Equals("SEM GTIN"))
                {
                    return true;
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }

    public static class IsbnValidator
    {
        public static bool IsValidIsbn(this string isbn)
        {
            if (!isbn.IsValidIsbnFormat())
            {
                return false;
            }

            if (!isbn.IsValidIsnb10Checksum())
            {
                return isbn.IsValidIsbn13Checksum();
            }

            return true;
        }

        public static bool IsValidIsbnFormat(this string isbn)
        {
            return new Regex("ISBN(-1(?:(0)|3))?:?\\x20+(?(1)(?(2)(?:(?=.{13}$)\\d{1,5}([ -])\\d{1,7}\\3\\d{1,6}\\3(?:\\d|x)$)|(?:(?=.{17}$)97(?:8|9)([ -])\\d{1,5}\\4\\d{1,7}\\4\\d{1,6}\\4\\d$))|(?(.{13}$)(?:\\d{1,5}([ -])\\d{1,7}\\5\\d{1,6}\\5(?:\\d|x)$)|(?:(?=.{17}$)97(?:8|9)([ -])\\d{1,5}\\6\\d{1,7}\\6\\d{1,6}\\6\\d$)))").IsMatch(isbn);
        }

        public static bool IsValidIsnb10Checksum(this string isbn)
        {
            if (isbn == null)
            {
                return false;
            }

            isbn = isbn.RemoveIsbn10();
            if (!ValidatorHelpers.IsDigitsOnly(isbn) || isbn.Length != 10)
            {
                return false;
            }

            int num = 0;
            for (int i = 0; i < 9; i++)
            {
                num += (i + 1) * int.Parse(isbn[i].ToString());
            }

            int num2 = num % 11;
            if (num2 != 10)
            {
                return isbn[9] == (ushort)(48 + num2);
            }

            return isbn[9] == 'X';
        }

        public static bool IsValidIsbn13Checksum(this string isbn)
        {
            isbn = isbn.RemoveIsbn13();
            if (!ValidatorHelpers.IsDigitsOnly(isbn) || isbn.Length != 13)
            {
                return false;
            }

            int num = 1;
            int num2 = 0;
            for (int i = 0; i < isbn.Length - 1; i++)
            {
                num2 += int.Parse(isbn[i].ToString()) * num;
                num = (num + 2) % 4;
            }

            int num3 = (10 - num2 % 10) % 10;
            return isbn[isbn.Length - 1] == (ushort)(48 + num3);
        }

        private static string RemoveIsbn13(this string isbn)
        {
            isbn = new Regex("ISBN(?:-13)?:?\\x20*").Replace(isbn, "");
            return isbn.Replace("-", "").Replace(" ", "");
        }

        private static string RemoveIsbn10(this string isbn)
        {
            isbn = new Regex("ISBN(?:-10)?:?\\x20*").Replace(isbn, "");
            return isbn.Replace("-", "").Replace(" ", "");
        }

        public static bool EanIsIsbn(this string ean)
        {
            return new Regex("^(97(8|9))?\\d{9}(\\d|X)$").IsMatch(ean);
        }
    }

    public static class UpcValidator
    {
        public static bool IsValidUpc(this string upc)
        {
            try
            {
                return ValidatorHelpers.IsValidGtin(upc);
            }
            catch
            {
                return false;
            }
        }
    }

    internal class ValidatorHelpers
    {
        public static bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                {
                    return false;
                }
            }

            return true;
        }

        public static bool IsStandardChecksumValid(string val)
        {
            int num = 1;
            int num2 = 0;
            for (int i = 0; i < val.Length - 1; i++)
            {
                num2 += int.Parse(val[i].ToString()) * num;
                num = (num + 2) % 4;
            }

            int num3 = (10 - num2 % 10) % 10;
            return val[val.Length - 1] == (ushort)(48 + num3);
        }

        public static bool IsValidGtin(string code)
        {
            if (code != new Regex("[^0-9]").Replace(code, ""))
            {
                return false;
            }

            switch (code.Length)
            {
                case 8:
                    code = "000000" + code;
                    goto case 14;
                case 12:
                    code = "00" + code;
                    goto case 14;
                case 13:
                    code = "0" + code;
                    goto case 14;
                case 14:
                    {
                        int[] array = new int[13]
                        {
                    int.Parse(code[0].ToString()) * 3,
                    int.Parse(code[1].ToString()),
                    int.Parse(code[2].ToString()) * 3,
                    int.Parse(code[3].ToString()),
                    int.Parse(code[4].ToString()) * 3,
                    int.Parse(code[5].ToString()),
                    int.Parse(code[6].ToString()) * 3,
                    int.Parse(code[7].ToString()),
                    int.Parse(code[8].ToString()) * 3,
                    int.Parse(code[9].ToString()),
                    int.Parse(code[10].ToString()) * 3,
                    int.Parse(code[11].ToString()),
                    int.Parse(code[12].ToString()) * 3
                        };
                        return (10 - (array[0] + array[1] + array[2] + array[3] + array[4] + array[5] + array[6] + array[7] + array[8] + array[9] + array[10] + array[11] + array[12]) % 10) % 10 == int.Parse(code[13].ToString());
                    }
                default:
                    return false;
            }
        }
    }
}