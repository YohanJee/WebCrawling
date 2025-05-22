using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using Dapper;

class Program
{
    private static string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["sam"].ToString();

    static void Main()
    {
        // Load data from the database
        var products = LoadProducts();

        // Process each product
        foreach (var product in products)
        {
            Console.WriteLine($"Processing masterid: {product.masterid} with Description: {product.Dsc}");
            var BestMatch1 = FindBestMatch(product.Dsc);
            if (BestMatch1 != null)
            {
                // Update the database with the best match
                UpdateProductBestMatch(product.masterid, BestMatch1.FullName, BestMatch1.Rank);
                Console.WriteLine($"Updated masterid: {product.masterid} with BestMatch1: {BestMatch1.FullName}, Rank: {BestMatch1.Rank}");
            }
            else
            {
                Console.WriteLine($"No match found for masterid: {product.masterid}");
            }
        }
    }

    private static List<Product> LoadProducts()
    {
        using (var connection = new SqlConnection(connectionString))
        {
            return connection.Query<Product>("SELECT masterid, Dsc FROM dbo.Product").ToList();
        }
    }

    private static MatchResult FindBestMatch(string description)
    {
        using (var connection = new SqlConnection(connectionString))
        {
            // Clean and format the search string to handle special characters and reserved words properly
            var formattedDescription = CleanAndFormatSearchString(description);

            var query = @"
            SELECT TOP 1 c.FullName, ct.RANK
            FROM dbo.Category c
            INNER JOIN CONTAINSTABLE(dbo.Category, FullName, @Description) AS ct ON c.FNID = ct.[KEY]
            ORDER BY ct.RANK DESC";

            return connection.QuerySingleOrDefault<MatchResult>(query, new { Description = formattedDescription });
        }
    }








    private static string CleanAndFormatSearchString(string input)
    {
        // List of SQL Server Full-Text Search reserved words
        var reservedWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "AND", "OR", "NOT", "NEAR"
    };

        // Remove special characters that are likely to cause syntax errors
        var cleanedString = new string(input.Where(c => char.IsLetterOrDigit(c) || char.IsWhiteSpace(c)).ToArray());

        // Split the cleaned string into individual words
        var words = cleanedString.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        // Enclose each word in double quotes and join with OR
        var formattedString = string.Join(" OR ", words.Select(word => $"\"{word}*\""));

        return formattedString;
    }



    private static void UpdateProductBestMatch(int masterid, string BestMatch1, int rank)
    {
        using (var connection = new SqlConnection(connectionString))
        {
            connection.Execute("UPDATE dbo.Product SET BestMatch1 = @BestMatch1, BestMatchRank1 = @Rank WHERE masterid = @masterid",
                new { BestMatch1 = BestMatch1, Rank = rank, masterid = masterid });
        }
    }

    class Product
    {
        public int masterid { get; set; }
        public string Dsc { get; set; }
    }

    class MatchResult
    {
        public string FullName { get; set; }
        public int Rank { get; set; }
    }
}
