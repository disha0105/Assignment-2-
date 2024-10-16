using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string filePath = @"C:\Book.xlsx";

        if (!File.Exists(filePath))
        {
            Console.WriteLine("The specified file does not exist.");
            return;
        }

        try
        {
            using var workbook = new XLWorkbook(filePath);
            var backupData = workbook.Worksheet("BackUP Data ");
            var currentData = workbook.Worksheet("Current Data ");

            // Dictionary to store backup passwords
            Dictionary<string, string> dictionary = new Dictionary<string, string>();

            // Read backup data
            int rowNum = 2; // Assuming the data starts from row 2
            while (!backupData.Cell("A" + rowNum).IsEmpty())
            {
                string userId = backupData.Cell("A" + rowNum).GetString(); // Get user ID
                string passwordHash = backupData.Cell("C" + rowNum).GetString(); // Get password hash

                dictionary.Add(userId, passwordHash);

                rowNum++;
            }

            // Reset rowNum for current data
            rowNum = 2;
            int cnt = 0;

            // Compare current data with backup data
            while (!currentData.Cell("A" + rowNum).IsEmpty())
            {
                string userId = currentData.Cell("A" + rowNum).GetString(); // Get user ID
                string passwordHash = currentData.Cell("C" + rowNum).GetString(); // Get new password hash
               

                if (dictionary.ContainsKey(userId) && dictionary[userId] != passwordHash)
                {
                    Console.WriteLine($"Row {rowNum}: User {currentData.Cell("D" + rowNum).GetString()}");
                    Console.WriteLine($"Old Password Hash: {dictionary[userId]}");
                    Console.WriteLine($"New Password Hash: {passwordHash}");
                    Console.WriteLine("--------------------------------------------------");
                    cnt++;
                }

                rowNum++;
            }

            // Output the count of changed passwords
            Console.WriteLine($"Number of different users with changed passwords: {cnt}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
