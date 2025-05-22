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
            var BestMatch = FindBestMatch(product.Dsc);
            if (BestMatch != null)
            {
                // Update the database with the best match
                UpdateProductBestMatch(product.masterid, BestMatch.FullName, BestMatch.Rank);
                Console.WriteLine($"Updated masterid: {product.masterid} with BestMatch: {BestMatch.FullName}, Rank: {BestMatch.Rank}");
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
            var query = @"
                SELECT TOP 1 c.FullName, ft.RANK
                FROM dbo.Category c
                INNER JOIN FREETEXTTABLE(dbo.Category, FullName, @Description) AS ft ON c.FNID = ft.[KEY]
                ORDER BY ft.RANK DESC";

            return connection.QuerySingleOrDefault<MatchResult>(query, new { Description = description });
        }
    }

    private static void UpdateProductBestMatch(int masterid, string BestMatch, int rank)
    {
        using (var connection = new SqlConnection(connectionString))
        {
            connection.Execute("UPDATE dbo.Product SET BestMatch = @BestMatch, BestMatchRank = @Rank WHERE masterid = @masterid",
                new { BestMatch = BestMatch, Rank = rank, masterid = masterid });
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
