using System;
using System.Collections.Generic;
using System.IO;

class Program
{
    static void Main()
    {
        Dictionary<string, decimal> inventory = new Dictionary<string, decimal>();

        while (true)
        {
            Console.Write("Enter item name (type 'done' to finish): ");
            string itemName = Console.ReadLine();

            if (itemName.ToLower() == "done")
                break;

            if (string.IsNullOrWhiteSpace(itemName))
            {
                Console.WriteLine("Item name cannot be empty. Please try again.");
                continue;
            }

            decimal itemValue;
            while (true)
            {
                Console.Write($"Enter value for {itemName}: ");
                string valueInput = Console.ReadLine();

                if (decimal.TryParse(valueInput, out itemValue) && itemValue >= 0)
                    break;
                Console.WriteLine("Invalid value. Please enter a non-negative numeric value.");
            }

            inventory[itemName] = itemValue;
        }

        // Display inventory summary
        Console.WriteLine("\n------ Inventory Summary ------");
        decimal totalValue = 0m;
        foreach (var item in inventory)
        {
            Console.WriteLine($"{item.Key}: ${item.Value:F2}");
            totalValue += item.Value;
        }

        Console.WriteLine($"Total value of inventory: ${totalValue:F2}");

        // Ask if the user wants to save to file
        Console.Write("Do you want to save the inventory to a file? (yes/no): ");
        string saveChoice = Console.ReadLine().ToLower();

        if (saveChoice == "yes" || saveChoice == "y")
        {
            string filename = @"C:\Users\Andre\Downloads\" + $"inventory_value{DateTime.Now:yyyyMMdd_HHmmss}.txt";

            try
            {
                using (StreamWriter writer = new StreamWriter(filename))
                {
                    foreach (var item in inventory)
                    {
                        writer.WriteLine($"{item.Key},{item.Value:F2}");
                    }
                    writer.WriteLine($"Total,{totalValue:F2}");
                }
                Console.WriteLine($"\nInventory saved to {filename}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError saving file: {ex.Message}");
            }
        }
    }
}