using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace LicenseGenerator
{
    internal static class Program
    {
        private const string Alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
        private static readonly Regex LicenseRegex = new Regex("[A-Z0-9]{3}-[A-Z0-9]{3}-[A-Z0-9]{3}-[A-Z0-9]{4}", RegexOptions.Compiled);

        private static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("Inventory Compare - License Generator");
            Console.WriteLine("-------------------------------------");

            int count = ReadCount(args);
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "licenses.json");

            List<string> licenses = ReadExistingLicenses(outputPath);
            HashSet<string> uniqueLicenses = new HashSet<string>(licenses);
            List<string> createdLicenses = new List<string>();

            while (createdLicenses.Count < count)
            {
                string license = CreateLicense();
                if (uniqueLicenses.Add(license))
                {
                    createdLicenses.Add(license);
                }
            }

            WriteLicenseFile(outputPath, uniqueLicenses.OrderBy(value => value).ToList());
            File.SetAttributes(outputPath, File.GetAttributes(outputPath) | FileAttributes.Hidden);

            Console.WriteLine();
            Console.WriteLine("License mới:");
            foreach (string license in createdLicenses)
            {
                Console.WriteLine(license);
            }

            Console.WriteLine();
            Console.WriteLine("Đã cập nhật file ẩn:");
            Console.WriteLine(outputPath);
            Console.WriteLine();
            Console.WriteLine("Nhấn phím bất kỳ để đóng.");
            Console.ReadKey(true);
            return 0;
        }

        private static int ReadCount(string[] args)
        {
            int argCount;
            if (args.Length > 0 && int.TryParse(args[0], out argCount) && argCount > 0)
            {
                return argCount;
            }

            Console.Write("Số lượng license cần tạo: ");
            string input = Console.ReadLine();
            int count;
            if (int.TryParse(input, out count) && count > 0)
            {
                return count;
            }

            return 1;
        }

        private static List<string> ReadExistingLicenses(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return new List<string>();
            }

            string content = File.ReadAllText(filePath).ToUpperInvariant();
            return LicenseRegex.Matches(content)
                .Cast<Match>()
                .Select(match => match.Value)
                .Distinct()
                .ToList();
        }

        private static void WriteLicenseFile(string filePath, List<string> licenses)
        {
            string json = "{\r\n  \"licenses\": [\r\n" +
                string.Join(",\r\n", licenses.Select(license => "    \"" + license + "\"")) +
                "\r\n  ]\r\n}\r\n";
            File.WriteAllText(filePath, json);
        }

        private static string CreateLicense()
        {
            return CreateSegment(3) + "-" + CreateSegment(3) + "-" + CreateSegment(3) + "-" + CreateSegment(4);
        }

        private static string CreateSegment(int length)
        {
            char[] chars = new char[length];
            byte[] bytes = new byte[length];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }

            for (int index = 0; index < length; index++)
            {
                chars[index] = Alphabet[bytes[index] % Alphabet.Length];
            }

            return new string(chars);
        }
    }
}
