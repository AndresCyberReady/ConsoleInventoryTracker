using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace InventoryValue
{
    public partial class MainForm : Form
    {
        private DataTable inventoryTable;
        private DataGridView dataGridView;
        private Label totalLabel;
        private Button importExcelButton;
        private Button addItemButton;
        private Button saveButton;
        private Button deleteButton;
        private Button reportButton;
        private System.Windows.Forms.Timer totalUpdateTimer;

        public MainForm()
        {
            InitializeComponent();
            InitializeInventory();
        }

        private void InitializeComponent()
        {
            this.Text = "🏠 Home Inventory Manager";
            this.Size = new Size(900, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(700, 500);
            this.BackColor = Color.FromArgb(245, 247, 250);
            this.ForeColor = Color.FromArgb(33, 37, 41);

            // Create DataTable for inventory
            inventoryTable = new DataTable();
            inventoryTable.Columns.Add("Item Name", typeof(string));
            inventoryTable.Columns.Add("Category", typeof(string));
            inventoryTable.Columns.Add("Value", typeof(decimal));
            inventoryTable.Columns.Add("Year Purchased", typeof(int));

            // Create header panel
            Panel headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = Color.FromArgb(52, 152, 219),
                Padding = new Padding(20, 15, 20, 15)
            };

            Label titleLabel = new Label
            {
                Text = "🏠 Home Inventory Manager",
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                ForeColor = Color.White,
                Dock = DockStyle.Left,
                AutoSize = true
            };

            Label subtitleLabel = new Label
            {
                Text = "Track your valuable items and their worth",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(236, 240, 241),
                Dock = DockStyle.Bottom,
                Height = 20,
                Padding = new Padding(0, 0, 0, 5)
            };

            headerPanel.Controls.Add(subtitleLabel);
            headerPanel.Controls.Add(titleLabel);

            // Create DataGridView with modern styling
            dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AllowUserToAddRows = false,
                ReadOnly = false,
                DataSource = inventoryTable,
                Margin = new Padding(15),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor = Color.FromArgb(230, 230, 230),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Font = new Font("Segoe UI", 9),
                    Padding = new Padding(5, 3, 5, 3),
                    SelectionBackColor = Color.FromArgb(52, 152, 219),
                    SelectionForeColor = Color.White
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(249, 249, 249)
                }
            };
            dataGridView.CellValueChanged += DataGridView_CellValueChanged;
            dataGridView.UserDeletingRow += DataGridView_UserDeletingRow;
            dataGridView.DataBindingComplete += DataGridView_DataBindingComplete;
            dataGridView.VisibleChanged += DataGridView_VisibleChanged;
            dataGridView.ColumnHeaderMouseDoubleClick += DataGridView_ColumnHeaderMouseDoubleClick;
            this.Resize += MainForm_Resize;
            this.Load += MainForm_Load;
            this.Shown += MainForm_Shown;

            // Style DataGridView headers
            dataGridView.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(44, 62, 80),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Padding = new Padding(5),
                Alignment = DataGridViewContentAlignment.MiddleLeft
            };
            dataGridView.EnableHeadersVisualStyles = false;

            // Create buttons panel with gradient background
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 70,
                BackColor = Color.FromArgb(236, 240, 241),
                Padding = new Padding(20, 15, 20, 15),
                AutoScroll = true
            };

            importExcelButton = CreateStyledButton("📊 Import Excel", new Size(140, 40), Color.FromArgb(46, 204, 113));
            importExcelButton.Location = new Point(20, 15);
            importExcelButton.Click += ImportExcelButton_Click;

            addItemButton = CreateStyledButton("➕ Add Item", new Size(120, 40), Color.FromArgb(52, 152, 219));
            addItemButton.Location = new Point(170, 15);
            addItemButton.Click += AddItemButton_Click;

            deleteButton = CreateStyledButton("🗑️ Delete Item", new Size(130, 40), Color.FromArgb(231, 76, 60));
            deleteButton.Location = new Point(300, 15);
            deleteButton.Click += DeleteButton_Click;

            saveButton = CreateStyledButton("💾 Save to File", new Size(130, 40), Color.FromArgb(155, 89, 182));
            saveButton.Location = new Point(440, 15);
            saveButton.Click += SaveButton_Click;

            reportButton = CreateStyledButton("📋 Insurance Report", new Size(150, 40), Color.FromArgb(241, 196, 15));
            reportButton.Location = new Point(580, 15);
            reportButton.Click += ReportButton_Click;

            buttonPanel.Controls.Add(importExcelButton);
            buttonPanel.Controls.Add(addItemButton);
            buttonPanel.Controls.Add(deleteButton);
            buttonPanel.Controls.Add(saveButton);
            buttonPanel.Controls.Add(reportButton);

            // Create total label panel with accent color
            Panel totalPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.FromArgb(44, 62, 80),
                Padding = new Padding(20, 10, 20, 10)
            };

            Label totalTextLabel = new Label
            {
                Text = "Total Inventory Value:",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.FromArgb(236, 240, 241),
                Dock = DockStyle.Left,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft
            };

            totalLabel = new Label
            {
                Text = "$0.00",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = Color.FromArgb(46, 204, 113),
                Dock = DockStyle.Right,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(10, 0, 0, 0)
            };

            totalPanel.Controls.Add(totalTextLabel);
            totalPanel.Controls.Add(totalLabel);

            // Create main panel with padding
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15),
                BackColor = Color.FromArgb(245, 247, 250)
            };

            mainPanel.Controls.Add(dataGridView);

            // Add controls to form
            this.Controls.Add(mainPanel);
            this.Controls.Add(totalPanel);
            this.Controls.Add(buttonPanel);
            this.Controls.Add(headerPanel);

            UpdateTotal();

            // Initialize timer to update total every 2 seconds
            totalUpdateTimer = new System.Windows.Forms.Timer();
            totalUpdateTimer.Interval = 2000; // 2 seconds
            totalUpdateTimer.Tick += (s, args) => UpdateTotal();
            totalUpdateTimer.Start();

            // Format columns immediately after setup
            if (dataGridView.Columns.Count > 0)
            {
                ApplyColumnFormatting();
            }
            else
            {
                // If columns don't exist yet, use a timer
                System.Windows.Forms.Timer initTimer = new System.Windows.Forms.Timer();
                initTimer.Interval = 50;
                int attempts = 0;
                initTimer.Tick += (s, args) =>
                {
                    attempts++;
                    if (dataGridView.Columns.Count > 0 || attempts > 20)
                    {
                        initTimer.Stop();
                        initTimer.Dispose();
                        if (dataGridView.Columns.Count > 0)
                        {
                            ApplyColumnFormatting();
                        }
                    }
                };
                initTimer.Start();
            }

            // Dispose timer when form closes
            this.FormClosing += (s, e) =>
            {
                if (totalUpdateTimer != null)
                {
                    totalUpdateTimer.Stop();
                    totalUpdateTimer.Dispose();
                }
            };
        }

        private Button CreateStyledButton(string text, Size size, Color backColor)
        {
            Button button = new Button
            {
                Text = text,
                Size = size,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = backColor,
                Cursor = Cursors.Hand
            };

            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(
                Math.Min(255, backColor.R + 20),
                Math.Min(255, backColor.G + 20),
                Math.Min(255, backColor.B + 20)
            );
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(
                Math.Max(0, backColor.R - 20),
                Math.Max(0, backColor.G - 20),
                Math.Max(0, backColor.B - 20)
            );

            return button;
        }

        private void InitializeInventory()
        {
            // Initialize with empty inventory
            inventoryTable.Clear();
        }

        private void ImportExcelButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string extension = Path.GetExtension(openFileDialog.FileName).ToLower();
                        if (extension == ".xlsx")
                        {
                            ImportExcelFile(openFileDialog.FileName);
                            MessageBox.Show("Excel file imported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else if (extension == ".csv")
                        {
                            ImportCsvFile(openFileDialog.FileName);
                            MessageBox.Show("CSV file imported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            // Try Excel first, then CSV
                            try
                            {
                                ImportExcelFile(openFileDialog.FileName);
                                MessageBox.Show("File imported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch
                            {
                                ImportCsvFile(openFileDialog.FileName);
                                MessageBox.Show("File imported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error importing file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ImportExcelFile(string filePath)
        {
            // Set EPPlus license context (required for non-commercial use)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            List<string> errors = new List<string>();
            int importedCount = 0;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Get the first worksheet
                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet == null)
                {
                    throw new Exception("Excel file does not contain any worksheets.");
                }

                int rowCount = worksheet.Dimension?.Rows ?? 0;
                int colCount = worksheet.Dimension?.Columns ?? 0;

                if (rowCount == 0)
                {
                    throw new Exception("Excel file is empty.");
                }

                // Determine which columns contain item name, category, value, and year purchased
                int nameColumn = -1;
                int categoryColumn = -1;
                int valueColumn = -1;
                int yearColumn = -1;
                int startRow = 1;

                // Check first row for headers
                bool hasHeader = false;
                for (int col = 1; col <= colCount; col++)
                {
                    var cellValue = worksheet.Cells[1, col].Value?.ToString()?.ToLower() ?? "";
                    if (cellValue.Contains("name") || cellValue.Contains("item"))
                    {
                        nameColumn = col;
                        hasHeader = true;
                    }
                    if (cellValue.Contains("category") || cellValue.Contains("type") || cellValue.Contains("class"))
                    {
                        categoryColumn = col;
                        hasHeader = true;
                    }
                    if (cellValue.Contains("value") || cellValue.Contains("price") || cellValue.Contains("cost"))
                    {
                        valueColumn = col;
                        hasHeader = true;
                    }
                    if (cellValue.Contains("year") || cellValue.Contains("purchased") || cellValue.Contains("date"))
                    {
                        yearColumn = col;
                        hasHeader = true;
                    }
                }

                // If no header detected, assume: name, category, value, year
                if (!hasHeader)
                {
                    nameColumn = 1;
                    categoryColumn = 2;
                    valueColumn = 3;
                    yearColumn = 4;
                    startRow = 1;
                }
                else
                {
                    startRow = 2; // Skip header row
                }

                // Validate we found the columns - use defaults if not found
                if (nameColumn == -1) nameColumn = 1;
                if (categoryColumn == -1) categoryColumn = 2;
                if (valueColumn == -1) valueColumn = 3;
                if (yearColumn == -1) yearColumn = 4;

                // Process rows
                for (int row = startRow; row <= rowCount; row++)
                {
                    var nameCell = worksheet.Cells[row, nameColumn].Value;
                    var categoryCell = worksheet.Cells[row, categoryColumn].Value;
                    var valueCell = worksheet.Cells[row, valueColumn].Value;
                    var yearCell = worksheet.Cells[row, yearColumn].Value;

                    if (nameCell == null && categoryCell == null && valueCell == null && yearCell == null)
                        continue; // Skip empty rows

                    string itemName = nameCell?.ToString()?.Trim() ?? "";
                    string category = categoryCell?.ToString()?.Trim() ?? "";
                    string valueStr = valueCell?.ToString()?.Trim() ?? "";
                    string yearStr = yearCell?.ToString()?.Trim() ?? "";

                    // Skip total row
                    if (itemName.Equals("Total", StringComparison.OrdinalIgnoreCase))
                        continue;

            if (string.IsNullOrWhiteSpace(itemName))
                    {
                        errors.Add($"Row {row} skipped: Empty item name");
                        continue;
                    }

                    // Try to parse value - handle both string and numeric values
                    decimal value = 0m;
                    bool valueParsed = false;

                    if (valueCell != null)
                    {
                        if (valueCell is double || valueCell is decimal || valueCell is int || valueCell is long)
                        {
                            value = Convert.ToDecimal(valueCell);
                            valueParsed = true;
                        }
                        else
                        {
                            valueParsed = decimal.TryParse(valueStr, out value);
                        }
                    }

                    // Try to parse year - handle both string and numeric values, and dates
                    int year = 0;
                    bool yearParsed = false;

                    if (yearCell != null)
                    {
                        if (yearCell is DateTime dateTime)
                        {
                            year = dateTime.Year;
                            yearParsed = true;
                        }
                        else if (yearCell is double || yearCell is int || yearCell is long)
                        {
                            int yearInt = Convert.ToInt32(yearCell);
                            // Validate year is reasonable (between 1900 and current year + 1)
                            if (yearInt >= 1900 && yearInt <= DateTime.Now.Year + 1)
                            {
                                year = yearInt;
                                yearParsed = true;
                            }
                        }
                        else
                        {
                            if (int.TryParse(yearStr, out int yearInt) && yearInt >= 1900 && yearInt <= DateTime.Now.Year + 1)
                            {
                                year = yearInt;
                                yearParsed = true;
                            }
                        }
                    }

                    if (valueParsed && value >= 0)
                    {
                        // Check if item already exists
                        DataRow? existingRow = inventoryTable.AsEnumerable()
                            .FirstOrDefault(r => r.Field<string>("Item Name")?.Equals(itemName, StringComparison.OrdinalIgnoreCase) == true);

                        if (existingRow != null)
                        {
                            // Update existing item
                            existingRow["Category"] = string.IsNullOrWhiteSpace(category) ? DBNull.Value : category;
                            existingRow["Value"] = value;
                            if (yearParsed)
                            {
                                existingRow["Year Purchased"] = year;
                            }
                            else
                            {
                                existingRow["Year Purchased"] = DBNull.Value;
                            }
                        }
                        else
                        {
                            // Add new item
                            if (yearParsed)
                            {
                                inventoryTable.Rows.Add(itemName, string.IsNullOrWhiteSpace(category) ? DBNull.Value : category, value, year);
                            }
                            else
                            {
                                inventoryTable.Rows.Add(itemName, string.IsNullOrWhiteSpace(category) ? DBNull.Value : category, value, DBNull.Value);
                            }
                        }
                        importedCount++;
                    }
                    else
                    {
                        errors.Add($"Row {row} skipped: Invalid value '{valueStr}' for item '{itemName}'");
                    }
                }
            }

            UpdateTotal();

            if (errors.Count > 0)
            {
                string errorMessage = $"Imported {importedCount} items successfully.\n\n" +
                                     $"Warnings ({errors.Count}):\n" +
                                     string.Join("\n", errors.Take(10));
                if (errors.Count > 10)
                {
                    errorMessage += $"\n... and {errors.Count - 10} more warnings.";
                }
                MessageBox.Show(errorMessage, "Import Complete with Warnings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void AddItemButton_Click(object sender, EventArgs e)
        {
            using (AddItemDialog dialog = new AddItemDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string itemName = dialog.ItemName;
                    decimal value = dialog.ItemValue;
                    int? year = dialog.YearPurchased;

                    // Check if item already exists
                    DataRow? existingRow = inventoryTable.AsEnumerable()
                        .FirstOrDefault(row => row.Field<string>("Item Name")?.Equals(itemName, StringComparison.OrdinalIgnoreCase) == true);

                    if (existingRow != null)
                    {
                        var result = MessageBox.Show(
                            $"Item '{itemName}' already exists. Do you want to update its value?",
                            "Item Exists",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            existingRow["Category"] = string.IsNullOrWhiteSpace(dialog.Category) ? DBNull.Value : dialog.Category;
                            existingRow["Value"] = value;
                            if (year.HasValue)
                            {
                                existingRow["Year Purchased"] = year.Value;
                            }
                            else
                            {
                                existingRow["Year Purchased"] = DBNull.Value;
                            }
                        }
                    }
                    else
                    {
                        object category = string.IsNullOrWhiteSpace(dialog.Category) ? DBNull.Value : dialog.Category;
                        if (year.HasValue)
                        {
                            inventoryTable.Rows.Add(itemName, category, value, year.Value);
                        }
                        else
                        {
                            inventoryTable.Rows.Add(itemName, category, value, DBNull.Value);
                        }
                    }

                    UpdateTotal();
                }
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*";
                saveFileDialog.FilterIndex = 1; // Default to Excel
                saveFileDialog.FileName = $"inventory_value{DateTime.Now:yyyyMMdd_HHmmss}";
                saveFileDialog.RestoreDirectory = true;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string extension = Path.GetExtension(saveFileDialog.FileName).ToLower();
                        decimal totalValue = CalculateTotal();

                        if (extension == ".xlsx")
                        {
                            SaveToExcel(saveFileDialog.FileName, totalValue);
                        }
                        else
                        {
                            // Save as CSV or text
                            using (StreamWriter writer = new StreamWriter(saveFileDialog.FileName))
                            {
                                writer.WriteLine("Item Name,Category,Value,Year Purchased");
                                foreach (DataRow row in inventoryTable.Rows)
                                {
                                    string categoryStr = row["Category"] == DBNull.Value ? "" : row["Category"].ToString();
                                    string yearStr = row["Year Purchased"] == DBNull.Value ? "" : row["Year Purchased"].ToString();
                                    writer.WriteLine($"{row["Item Name"]},{categoryStr},{row["Value"]:F2},{yearStr}");
                                }
                                writer.WriteLine($"Total,,{totalValue:F2},");
                            }
                        }
                        MessageBox.Show($"Inventory saved to {saveFileDialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void SaveToExcel(string filePath, decimal totalValue)
        {
            // Set EPPlus license context (required for non-commercial use)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Inventory");

                // Add headers
                worksheet.Cells[1, 1].Value = "Item Name";
                worksheet.Cells[1, 2].Value = "Category";
                worksheet.Cells[1, 3].Value = "Value";
                worksheet.Cells[1, 4].Value = "Year Purchased";

                // Style header row
                using (var range = worksheet.Cells[1, 1, 1, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(44, 62, 80));
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                }

                // Add data rows
                int row = 2;
                foreach (DataRow dataRow in inventoryTable.Rows)
                {
                    worksheet.Cells[row, 1].Value = dataRow["Item Name"]?.ToString() ?? "";
                    worksheet.Cells[row, 2].Value = dataRow["Category"] == DBNull.Value ? "" : dataRow["Category"].ToString();
                    worksheet.Cells[row, 3].Value = dataRow["Value"] == DBNull.Value ? 0 : Convert.ToDecimal(dataRow["Value"]);
                    worksheet.Cells[row, 3].Style.Numberformat.Format = "$#,##0.00";
                    
                    if (dataRow["Year Purchased"] != DBNull.Value)
                    {
                        worksheet.Cells[row, 4].Value = Convert.ToInt32(dataRow["Year Purchased"]);
                    }
                    row++;
                }

                // Add total row
                worksheet.Cells[row, 1].Value = "Total";
                worksheet.Cells[row, 3].Value = totalValue;
                worksheet.Cells[row, 3].Style.Numberformat.Format = "$#,##0.00";
                worksheet.Cells[row, 1, row, 3].Style.Font.Bold = true;

                // Auto-fit columns
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                package.SaveAs(new FileInfo(filePath));
            }
        }

        private void ImportCsvFile(string filePath)
        {
            List<string> errors = new List<string>();
            int importedCount = 0;

            using (StreamReader reader = new StreamReader(filePath))
            {
                string? line;
                bool isFirstLine = true;

                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // Skip header row if it exists
                    if (isFirstLine)
                    {
                        isFirstLine = false;
                        string lowerLine = line.ToLower();
                        if (lowerLine.Contains("name") || lowerLine.Contains("item") || 
                            lowerLine.Contains("category") || lowerLine.Contains("value") || 
                            lowerLine.Contains("year") || lowerLine.Contains("total"))
                        {
                            continue; // Skip header row
                        }
                    }

                    // Skip total row
                    if (line.ToLower().StartsWith("total,"))
                        continue;

                    // Parse CSV line
                    string[] parts = ParseCsvLine(line);

                    if (parts.Length >= 3)
                    {
                        string itemName = parts[0].Trim();
                        string category = parts.Length > 1 ? parts[1].Trim() : "";
                        string valueStr = parts.Length > 2 ? parts[2].Trim() : "";
                        string yearStr = parts.Length > 3 ? parts[3].Trim() : "";

                        if (string.IsNullOrWhiteSpace(itemName))
                        {
                            errors.Add($"Line skipped: Empty item name");
                continue;
                        }

                        if (decimal.TryParse(valueStr, out decimal value) && value >= 0)
                        {
                            // Parse year if provided
                            int? year = null;
                            if (!string.IsNullOrWhiteSpace(yearStr) && yearStr != "N/A")
                            {
                                if (int.TryParse(yearStr, out int yearInt) && yearInt >= 1900 && yearInt <= DateTime.Now.Year + 1)
                                {
                                    year = yearInt;
                                }
                            }

                            // Check if item already exists
                            DataRow? existingRow = inventoryTable.AsEnumerable()
                                .FirstOrDefault(r => r.Field<string>("Item Name")?.Equals(itemName, StringComparison.OrdinalIgnoreCase) == true);

                            if (existingRow != null)
                            {
                                // Update existing item
                                existingRow["Category"] = string.IsNullOrWhiteSpace(category) ? DBNull.Value : category;
                                existingRow["Value"] = value;
                                existingRow["Year Purchased"] = year.HasValue ? (object)year.Value : DBNull.Value;
                            }
                            else
                            {
                                // Add new item
                                inventoryTable.Rows.Add(
                                    itemName,
                                    string.IsNullOrWhiteSpace(category) ? DBNull.Value : category,
                                    value,
                                    year.HasValue ? (object)year.Value : DBNull.Value
                                );
                            }
                            importedCount++;
                        }
                        else
                        {
                            errors.Add($"Line skipped: Invalid value '{valueStr}' for item '{itemName}'");
                        }
                    }
                    else
                    {
                        errors.Add($"Line skipped: Invalid format - '{line}'");
                    }
                }
            }

            UpdateTotal();

            if (errors.Count > 0)
            {
                string errorMessage = $"Imported {importedCount} items successfully.\n\n" +
                                     $"Warnings ({errors.Count}):\n" +
                                     string.Join("\n", errors.Take(10));
                if (errors.Count > 10)
                {
                    errorMessage += $"\n... and {errors.Count - 10} more warnings.";
                }
                MessageBox.Show(errorMessage, "Import Complete with Warnings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private string[] ParseCsvLine(string line)
        {
            List<string> fields = new List<string>();
            bool inQuotes = false;
            string currentField = "";

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(currentField);
                    currentField = "";
                }
                else
                {
                    currentField += c;
                }
            }

            fields.Add(currentField); // Add the last field
            return fields.ToArray();
        }

        private void DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                UpdateTotal();
            }
        }

        private void DataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            UpdateTotal();
        }

        private void DataGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            // Apply formatting immediately when data binding completes
            if (dataGridView.Columns.Count > 0)
            {
                ApplyColumnFormatting();
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Ensure columns are formatted when form loads
            UpdateColumnWidths();
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            // Set column widths after form is fully shown
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 10; // Small delay to ensure everything is ready
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                timer.Dispose();
                // Force column creation if needed
                if (dataGridView.Columns.Count == 0 && inventoryTable.Columns.Count > 0)
                {
                    // Trigger column creation by temporarily changing data source
                    var tempSource = dataGridView.DataSource;
                    dataGridView.DataSource = null;
                    dataGridView.DataSource = tempSource;
                }
                ApplyColumnFormatting();
            };
            timer.Start();
        }

        private void DataGridView_VisibleChanged(object sender, EventArgs e)
        {
            // Format columns when DataGridView becomes visible
            if (dataGridView.Visible)
            {
                UpdateColumnWidths();
            }
        }

        private void UpdateColumnWidths()
        {
            if (dataGridView == null) return;

            // Force columns to be created if they don't exist yet
            if (dataGridView.Columns.Count == 0 && inventoryTable.Columns.Count > 0)
            {
                // Columns should be auto-created, but if not, wait a bit
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 10;
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    if (dataGridView.Columns.Count > 0)
                    {
                        ApplyColumnFormatting();
                    }
                };
                timer.Start();
                return;
            }

            if (dataGridView.Columns.Count == 0) return;

            ApplyColumnFormatting();
        }

        private void ApplyColumnFormatting()
        {
            if (dataGridView == null || dataGridView.Columns.Count == 0) return;

            // Use form width if DataGridView width isn't available yet
            int availableWidth = dataGridView.Width > 100 ? dataGridView.Width - 20 : (this.Width > 100 ? this.Width - 100 : 800);

            // Set column widths and formatting - use fixed pixel widths for consistency
            if (dataGridView.Columns["Item Name"] != null)
            {
                dataGridView.Columns["Item Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGridView.Columns["Item Name"].Width = (int)(availableWidth * 0.35);
                dataGridView.Columns["Item Name"].MinimumWidth = 150;
                dataGridView.Columns["Item Name"].FillWeight = 35;
            }

            if (dataGridView.Columns["Category"] != null)
            {
                dataGridView.Columns["Category"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGridView.Columns["Category"].Width = (int)(availableWidth * 0.20);
                dataGridView.Columns["Category"].MinimumWidth = 100;
                dataGridView.Columns["Category"].FillWeight = 20;
            }

            if (dataGridView.Columns["Value"] != null)
            {
                dataGridView.Columns["Value"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGridView.Columns["Value"].DefaultCellStyle.Format = "C2";
                dataGridView.Columns["Value"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView.Columns["Value"].Width = (int)(availableWidth * 0.25);
                dataGridView.Columns["Value"].MinimumWidth = 100;
                dataGridView.Columns["Value"].FillWeight = 25;
            }

            if (dataGridView.Columns["Year Purchased"] != null)
            {
                dataGridView.Columns["Year Purchased"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGridView.Columns["Year Purchased"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView.Columns["Year Purchased"].Width = (int)(availableWidth * 0.20);
                dataGridView.Columns["Year Purchased"].MinimumWidth = 100;
                dataGridView.Columns["Year Purchased"].FillWeight = 20;
            }

            // Force a refresh to ensure columns are displayed
            dataGridView.Refresh();
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (dataGridView != null && dataGridView.Columns.Count > 0)
            {
                UpdateColumnWidths();
            }
        }

        private void DataGridView_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Auto-size all columns to fit their content, like Excel
            AutoSizeAllColumns();
        }

        private void AutoSizeAllColumns()
        {
            if (dataGridView == null || dataGridView.Columns.Count == 0) return;

            // Temporarily set AutoSizeMode to AllCells to calculate optimal widths
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            // Get the calculated widths
            int totalAutoSizedWidth = 0;
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                totalAutoSizedWidth += column.Width;
            }

            // Calculate available width
            int availableWidth = dataGridView.Width - (dataGridView.RowHeadersVisible ? dataGridView.RowHeadersWidth : 0) - 2;
            
            // If auto-sized columns fit within available width, keep them
            // Otherwise, proportionally scale them down
            if (totalAutoSizedWidth <= availableWidth)
            {
                // Keep the auto-sized widths, but set mode back to None to prevent future auto-sizing
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    int width = column.Width;
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    column.Width = width;
                }
            }
            else
            {
                // Scale down proportionally
                double scaleFactor = (double)availableWidth / totalAutoSizedWidth;
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    int width = (int)(column.Width * scaleFactor);
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    column.Width = Math.Max(width, column.MinimumWidth);
                }
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            if (dataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DataGridViewRow selectedRow = dataGridView.SelectedRows[0];
            string itemName = selectedRow.Cells["Item Name"].Value?.ToString() ?? "this item";

            var result = MessageBox.Show(
                $"Are you sure you want to delete '{itemName}'?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    dataGridView.Rows.Remove(selectedRow);
                    UpdateTotal();
            }
            catch (Exception ex)
                {
                    MessageBox.Show($"Error deleting item: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void UpdateTotal()
        {
            decimal total = CalculateTotal();
            totalLabel.Text = total.ToString("C2");
        }

        private decimal CalculateTotal()
        {
            decimal total = 0m;
            foreach (DataRow row in inventoryTable.Rows)
            {
                if (row["Value"] != DBNull.Value)
                {
                    total += Convert.ToDecimal(row["Value"]);
                }
            }
            return total;
        }

        private void ReportButton_Click(object sender, EventArgs e)
        {
            if (inventoryTable.Rows.Count == 0)
            {
                MessageBox.Show("No inventory items to analyze. Please add items first.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (InsuranceReportDialog reportDialog = new InsuranceReportDialog(inventoryTable))
            {
                reportDialog.ShowDialog();
            }
        }
    }

    // Dialog for adding new items
    public class AddItemDialog : Form
    {
        private TextBox itemNameTextBox;
        private ComboBox categoryComboBox;
        private TextBox valueTextBox;
        private TextBox yearTextBox;
        private Button okButton;
        private Button cancelButton;

        public string ItemName { get; private set; } = "";
        public string Category { get; private set; } = "";
        public decimal ItemValue { get; private set; } = 0m;
        public int? YearPurchased { get; private set; } = null;

        public AddItemDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "➕ Add New Item";
            this.Size = new Size(420, 320);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(245, 247, 250);

            // Header panel
            Panel headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = Color.FromArgb(52, 152, 219),
                Padding = new Padding(15, 10, 15, 10)
            };

            Label headerLabel = new Label
            {
                Text = "➕ Add New Item",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.White,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            headerPanel.Controls.Add(headerLabel);

            // Button panel
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.FromArgb(236, 240, 241),
                Padding = new Padding(15, 12, 15, 12)
            };

            okButton = new Button
            {
                Text = "✓ Add Item",
                DialogResult = DialogResult.OK,
                Location = new Point(200, 12),
                Size = new Size(90, 36),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(46, 204, 113),
                Cursor = Cursors.Hand
            };
            okButton.FlatAppearance.BorderSize = 0;
            okButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(39, 174, 96);
            okButton.Click += OkButton_Click;

            cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(300, 12),
                Size = new Size(80, 36),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(33, 37, 41),
                BackColor = Color.FromArgb(236, 240, 241),
                Cursor = Cursors.Hand
            };
            cancelButton.FlatAppearance.BorderSize = 0;
            cancelButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(189, 195, 199);

            buttonPanel.Controls.Add(okButton);
            buttonPanel.Controls.Add(cancelButton);

            // Main content panel
            Panel contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 25, 20, 20),
                BackColor = Color.FromArgb(245, 247, 250)
            };

            Label nameLabel = new Label
            {
                Text = "Item Name:",
                Location = new Point(15, 15),
                Size = new Size(110, 25),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 37, 41),
                TextAlign = ContentAlignment.MiddleLeft
            };

            itemNameTextBox = new TextBox
            {
                Location = new Point(130, 12),
                Size = new Size(240, 28),
                Font = new Font("Segoe UI", 9),
                BorderStyle = BorderStyle.FixedSingle
            };

            Label valueLabel = new Label
            {
                Text = "Value ($):",
                Location = new Point(15, 55),
                Size = new Size(110, 25),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 37, 41),
                TextAlign = ContentAlignment.MiddleLeft
            };

            valueTextBox = new TextBox
            {
                Location = new Point(130, 52),
                Size = new Size(240, 28),
                Font = new Font("Segoe UI", 9),
                BorderStyle = BorderStyle.FixedSingle
            };

            Label categoryLabel = new Label
            {
                Text = "Category:",
                Location = new Point(15, 95),
                Size = new Size(110, 25),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 37, 41),
                TextAlign = ContentAlignment.MiddleLeft
            };

            categoryComboBox = new ComboBox
            {
                Location = new Point(130, 92),
                Size = new Size(240, 28),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            
            // Add common categories
            categoryComboBox.Items.AddRange(new string[] {
                "Electronics", "Furniture", "Appliances", "Clothing", "Jewelry",
                "Tools", "Sports Equipment", "Books", "Art & Decor", "Kitchenware",
                "Vehicles", "Musical Instruments", "Collectibles", "Other"
            });

            Label yearLabel = new Label
            {
                Text = "Year Purchased:",
                Location = new Point(15, 135),
                Size = new Size(110, 25),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 37, 41),
                TextAlign = ContentAlignment.MiddleLeft
            };

            yearTextBox = new TextBox
            {
                Location = new Point(130, 132),
                Size = new Size(240, 28),
                Font = new Font("Segoe UI", 9),
                BorderStyle = BorderStyle.FixedSingle
            };

            contentPanel.Controls.Add(nameLabel);
            contentPanel.Controls.Add(itemNameTextBox);
            contentPanel.Controls.Add(valueLabel);
            contentPanel.Controls.Add(valueTextBox);
            contentPanel.Controls.Add(categoryLabel);
            contentPanel.Controls.Add(categoryComboBox);
            contentPanel.Controls.Add(yearLabel);
            contentPanel.Controls.Add(yearTextBox);

            this.Controls.Add(contentPanel);
            this.Controls.Add(buttonPanel);
            this.Controls.Add(headerPanel);

            this.AcceptButton = okButton;
            this.CancelButton = cancelButton;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(itemNameTextBox.Text))
            {
                MessageBox.Show("Item name cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.None;
                return;
            }

            if (!decimal.TryParse(valueTextBox.Text, out decimal value) || value < 0)
            {
                MessageBox.Show("Please enter a valid non-negative numeric value.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.None;
                return;
            }

            // Validate year if provided
            int? year = null;
            if (!string.IsNullOrWhiteSpace(yearTextBox.Text))
            {
                if (int.TryParse(yearTextBox.Text, out int yearInt))
                {
                    if (yearInt >= 1900 && yearInt <= DateTime.Now.Year + 1)
                    {
                        year = yearInt;
                    }
                    else
                    {
                        MessageBox.Show($"Please enter a valid year between 1900 and {DateTime.Now.Year + 1}.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        this.DialogResult = DialogResult.None;
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please enter a valid year (numeric value).", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.DialogResult = DialogResult.None;
                    return;
                }
            }

            ItemName = itemNameTextBox.Text.Trim();
            Category = categoryComboBox.Text.Trim();
            ItemValue = value;
            YearPurchased = year;
        }
    }

    // Insurance Report Dialog
    public class InsuranceReportDialog : Form
    {
        private DataTable inventoryTable;
        private DataGridView reportDataGridView;
        private Label summaryLabel;
        private Label topCategoryLabel;

        public InsuranceReportDialog(DataTable inventory)
        {
            inventoryTable = inventory;
            InitializeComponent();
            GenerateReport();
        }

        private void InitializeComponent()
        {
            this.Text = "📋 Insurance Report";
            this.Size = new Size(900, 700);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(245, 247, 250);
            this.MinimizeBox = false;
            this.MaximizeBox = false;

            // Header panel
            Panel headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = Color.FromArgb(241, 196, 15),
                Padding = new Padding(20, 15, 20, 15)
            };

            Label titleLabel = new Label
            {
                Text = "📋 Insurance Report",
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                ForeColor = Color.FromArgb(44, 62, 80),
                Dock = DockStyle.Left,
                AutoSize = true
            };

            Label subtitleLabel = new Label
            {
                Text = "Category Analysis for Insurance Purposes",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(44, 62, 80),
                Dock = DockStyle.Bottom,
                Height = 20,
                Padding = new Padding(0, 0, 0, 5)
            };

            headerPanel.Controls.Add(subtitleLabel);
            headerPanel.Controls.Add(titleLabel);

            // Summary panel
            Panel summaryPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100,
                BackColor = Color.FromArgb(44, 62, 80),
                Padding = new Padding(20, 15, 20, 15)
            };

            topCategoryLabel = new Label
            {
                Text = "",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.FromArgb(241, 196, 15),
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(0, 0, 0, 10)
            };

            summaryLabel = new Label
            {
                Text = "",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.FromArgb(236, 240, 241),
                Dock = DockStyle.Fill,
                AutoSize = true
            };

            summaryPanel.Controls.Add(summaryLabel);
            summaryPanel.Controls.Add(topCategoryLabel);

            // Data grid view for items
            reportDataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor = Color.FromArgb(230, 230, 230),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Font = new Font("Segoe UI", 9),
                    Padding = new Padding(5, 3, 5, 3)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(249, 249, 249)
                }
            };

            reportDataGridView.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(44, 62, 80),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Padding = new Padding(5),
                Alignment = DataGridViewContentAlignment.MiddleLeft
            };
            reportDataGridView.EnableHeadersVisualStyles = false;

            // Button panel
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.FromArgb(236, 240, 241),
                Padding = new Padding(20, 10, 20, 10)
            };

            Button closeButton = new Button
            {
                Text = "Close",
                DialogResult = DialogResult.OK,
                Size = new Size(100, 35),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(44, 62, 80),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Right | AnchorStyles.Top
            };
            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(52, 73, 94);
            closeButton.Location = new Point(buttonPanel.Width - 120, 12);

            buttonPanel.Controls.Add(closeButton);

            // Main panel
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15),
                BackColor = Color.FromArgb(245, 247, 250)
            };

            mainPanel.Controls.Add(reportDataGridView);

            this.Controls.Add(mainPanel);
            this.Controls.Add(buttonPanel);
            this.Controls.Add(summaryPanel);
            this.Controls.Add(headerPanel);

            this.AcceptButton = closeButton;
        }

        private void GenerateReport()
        {
            // Analyze categories
            var categoryAnalysis = new Dictionary<string, CategorySummary>();

            foreach (DataRow row in inventoryTable.Rows)
            {
                string category = row["Category"] == DBNull.Value || string.IsNullOrWhiteSpace(row["Category"]?.ToString()) 
                    ? "Uncategorized" 
                    : row["Category"].ToString();
                decimal value = row["Value"] == DBNull.Value ? 0m : Convert.ToDecimal(row["Value"]);
                string itemName = row["Item Name"]?.ToString() ?? "";

                if (!categoryAnalysis.ContainsKey(category))
                {
                    categoryAnalysis[category] = new CategorySummary
                    {
                        CategoryName = category,
                        TotalValue = 0m,
                        ItemCount = 0,
                        Items = new List<ItemInfo>()
                    };
                }

                categoryAnalysis[category].TotalValue += value;
                categoryAnalysis[category].ItemCount++;
                categoryAnalysis[category].Items.Add(new ItemInfo
                {
                    Name = itemName,
                    Value = value,
                    YearPurchased = row["Year Purchased"] == DBNull.Value ? null : (int?)Convert.ToInt32(row["Year Purchased"])
                });
            }

            // Find category with highest value
            var topCategory = categoryAnalysis.OrderByDescending(c => c.Value.TotalValue).FirstOrDefault();

            if (topCategory.Key != null)
            {
                topCategoryLabel.Text = $"🏆 Highest Value Category: {topCategory.Key}";
                
                string summary = $"Total Value: {topCategory.Value.TotalValue:C2} | " +
                               $"Items: {topCategory.Value.ItemCount} | " +
                               $"Average Value per Item: {(topCategory.Value.ItemCount > 0 ? topCategory.Value.TotalValue / topCategory.Value.ItemCount : 0):C2}";
                summaryLabel.Text = summary;

                // Create DataTable for the report
                DataTable reportTable = new DataTable();
                reportTable.Columns.Add("Item Name", typeof(string));
                reportTable.Columns.Add("Value", typeof(decimal));
                reportTable.Columns.Add("Year Purchased", typeof(string));

                foreach (var item in topCategory.Value.Items.OrderByDescending(i => i.Value))
                {
                    reportTable.Rows.Add(
                        item.Name,
                        item.Value,
                        item.YearPurchased?.ToString() ?? "N/A"
                    );
                }

                reportDataGridView.DataSource = reportTable;

                // Format columns
                if (reportDataGridView.Columns["Value"] != null)
                {
                    reportDataGridView.Columns["Value"].DefaultCellStyle.Format = "C2";
                    reportDataGridView.Columns["Value"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }

                if (reportDataGridView.Columns["Year Purchased"] != null)
                {
                    reportDataGridView.Columns["Year Purchased"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }
        }

        private class CategorySummary
        {
            public string CategoryName { get; set; } = "";
            public decimal TotalValue { get; set; }
            public int ItemCount { get; set; }
            public List<ItemInfo> Items { get; set; } = new List<ItemInfo>();
        }

        private class ItemInfo
        {
            public string Name { get; set; } = "";
            public decimal Value { get; set; }
            public int? YearPurchased { get; set; }
        }
    }

    // Main entry point
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
