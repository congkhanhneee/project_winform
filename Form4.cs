using MongoDB.Bson;
using MongoDB.Driver;
using ScottPlot;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form4 : Form
    {
        private readonly IMongoCollection<BsonDocument> identityCollection;
        private readonly IMongoCollection<BsonDocument> detailFileCollection;
        private readonly IMongoCollection<BsonDocument> actionCollection;
        private bool isDatabaseConnected = false;

        public Form4()
        {
            InitializeComponent();

            try
            {
                var client = new MongoClient("mongodb://localhost:27017");
                var database = client.GetDatabase("FileManagement");

                identityCollection = database.GetCollection<BsonDocument>("identity");
                detailFileCollection = database.GetCollection<BsonDocument>("detail_file");
                actionCollection = database.GetCollection<BsonDocument>("action");

                isDatabaseConnected = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể kết nối CSDL: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            if (!isDatabaseConnected) return;
            LoadUsers();
        }

        private void LoadUsers()
        {
            if (!isDatabaseConnected) return;

            var users = detailFileCollection.Distinct<string>("user", new BsonDocument()).ToList();
            cbUserFilter.Items.Clear();
            cbUserFilter.Items.Add("Tất cả");
            cbUserFilter.Items.AddRange(users.ToArray());
            cbUserFilter.SelectedIndex = 0;
        }

        private void btnLoadFileChart_Click(object sender, EventArgs e)
        {
            DrawFileTypeChartAsync();
        }

        private void btnLoadActionChart_Click(object sender, EventArgs e)
        {
            DrawActionChart();
        }

        private async Task DrawFileTypeChartAsync()
        {
            string directoryPath = @"D:\";

            if (!Directory.Exists(directoryPath))
            {
                MessageBox.Show("Thư mục không tồn tại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var fileCount = new ConcurrentDictionary<string, int>();
            int filesProcessed = 0;

            try
            {
                await Task.Run(() =>
                {
                    Debug.WriteLine("Bắt đầu quét thư mục...");
                    try
                    {
                        // Lấy danh sách thư mục trước, bỏ qua các thư mục hệ thống
                        var directories = GetAccessibleDirectories(directoryPath);

                        // Quét file từ các thư mục có thể truy cập
                        var files = directories.SelectMany(dir =>
                        {
                            try
                            {
                                return Directory.EnumerateFiles(dir, "*.*", SearchOption.TopDirectoryOnly)
                                    .Where(f => !IsSystemFile(f));
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Lỗi quét thư mục {dir}: {ex.Message}");
                                return Enumerable.Empty<string>();
                            }
                        }).Take(1000000);

                        Debug.WriteLine($"Tìm thấy {files.Count()} file để xử lý");

                        Parallel.ForEach(files, new ParallelOptions
                        {
                            MaxDegreeOfParallelism = Environment.ProcessorCount / 2
                        }, file =>
                        {
                            try
                            {
                                string ext = Path.GetExtension(file)?.ToLowerInvariant() ?? "";
                                if (!string.IsNullOrEmpty(ext))
                                {
                                    fileCount.AddOrUpdate(ext, 1, (_, old) => old + 1);
                                }
                                int processed = Interlocked.Increment(ref filesProcessed);
                                if (processed % 1000 == 0)
                                {
                                    Debug.WriteLine($"Đã xử lý: {processed} files");
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Lỗi xử lý file {file}: {ex.Message}");
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Lỗi trong quá trình quét: {ex.Message}");
                        throw;
                    }
                });
            }
            catch (AggregateException ae)
            {
                Debug.WriteLine("AggregateException xảy ra:");
                foreach (var ex in ae.Flatten().InnerExceptions)
                {
                    Debug.WriteLine($" - Lỗi chi tiết: {ex.Message}");
                }
                MessageBox.Show($"Lỗi tổng hợp: {ae.Message}\nXem Output Window để biết chi tiết.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show($"Không đủ quyền: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            catch (PathTooLongException ex)
            {
                MessageBox.Show($"Đường dẫn quá dài: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi không xác định: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (fileCount.IsEmpty)
            {
                MessageBox.Show("Không tìm thấy file nào!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Debug.WriteLine($"Hoàn tất quét: {filesProcessed} files, {fileCount.Count} loại file");
            await UpdateChartAsync(fileCount);
        }

        // Lấy danh sách thư mục có thể truy cập, bỏ qua thư mục hệ thống
        private IEnumerable<string> GetAccessibleDirectories(string rootPath)
        {
            var directories = new List<string> { rootPath };
            var systemFolders = new[]
            {
                "System Volume Information",
                "$RECYCLE.BIN",
                "Windows",
                "Program Files",
                "Program Files (x86)"
            };

            var queue = new Queue<string>();
            queue.Enqueue(rootPath);

            while (queue.Count > 0)
            {
                var currentDir = queue.Dequeue();
                try
                {
                    var subDirs = Directory.EnumerateDirectories(currentDir);
                    foreach (var subDir in subDirs)
                    {
                        string dirName = Path.GetFileName(subDir); // Lấy tên thư mục
                        if (!systemFolders.Contains(dirName, StringComparer.OrdinalIgnoreCase) && !IsSystemFolder(subDir))
                        {
                            directories.Add(subDir);
                            queue.Enqueue(subDir);
                        }
                        else
                        {
                            Debug.WriteLine($"Bỏ qua thư mục hệ thống: {subDir}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Không thể truy cập thư mục {currentDir}: {ex.Message}");
                }
            }

            return directories;
        }

        private bool IsSystemFile(string path)
        {
            try
            {
                var attributes = File.GetAttributes(path);
                return (attributes & (FileAttributes.System | FileAttributes.Hidden)) != 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi kiểm tra thuộc tính file {path}: {ex.Message}");
                return true;
            }
        }

        private bool IsSystemFolder(string path)
        {
            try
            {
                var attributes = new DirectoryInfo(path).Attributes;
                return (attributes & (FileAttributes.System | FileAttributes.Hidden)) != 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi kiểm tra thuộc tính thư mục {path}: {ex.Message}");
                return true;
            }
        }

        private async Task UpdateChartAsync(ConcurrentDictionary<string, int> fileCount)
        {
            await Task.Run(() =>
            {
                Debug.WriteLine("Bắt đầu vẽ biểu đồ...");
                try
                {
                    var sortedFiles = fileCount.OrderByDescending(kv => kv.Value).Take(20);
                    double[] values = sortedFiles.Select(kv => (double)kv.Value).ToArray();
                    string[] labels = sortedFiles.Select(kv => kv.Key).ToArray();
                    double[] positions = Enumerable.Range(0, values.Length).Select(x => x * 0.5).ToArray();

                    formsPlot1.Invoke((Action)(() =>
                    {
                        try
                        {
                            formsPlot1.Plot.Clear();

                            var bars = new List<ScottPlot.Bar>();
                            for (int i = 0; i < values.Length; i++)
                            {
                                bars.Add(new ScottPlot.Bar
                                {
                                    Position = positions[i],
                                    Value = values[i],
                                    FillColor = ScottPlot.Colors.Blue,
                                    Size = 0.4
                                });
                            }

                            formsPlot1.Plot.Add.Bars(bars);
                            formsPlot1.Plot.Axes.Bottom.SetTicks(positions, labels);
                            formsPlot1.Plot.Axes.SetLimitsX(positions.Min() - 0.2, positions.Max() + 0.2);
                            formsPlot1.Plot.Axes.SetLimitsY(0, values.Max() * 1.1);

                            formsPlot1.Plot.Title("Phân bố loại file (20 loại phổ biến)");
                            formsPlot1.Plot.XLabel("Loại file");
                            formsPlot1.Plot.YLabel("Số lượng");

                            formsPlot1.Refresh();
                            Debug.WriteLine("Hoàn tất vẽ biểu đồ");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Lỗi trong Invoke: {ex.Message}");
                            throw;
                        }
                    }));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Lỗi vẽ biểu đồ: {ex.Message}");
                    throw;
                }
            });
        }


        private void DrawActionChart()
        {
            if (!isDatabaseConnected) return;

            string selectedUser = cbUserFilter.SelectedItem.ToString();
            FilterDefinition<BsonDocument> filter = Builders<BsonDocument>.Filter.Empty; // No filter if "All" is selected

            if (selectedUser != "Tất cả")
            {
                filter = Builders<BsonDocument>.Filter.Eq("user", selectedUser);
            }

            // Get action counts from the database
            var actions = detailFileCollection.Find(filter).ToList()
                .Where(doc => doc.Contains("id_action"))  // Kiểm tra tồn tại
                .GroupBy(doc => doc["id_action"].IsInt32 ? doc["id_action"].AsInt32 : int.Parse(doc["id_action"].AsString))
                .ToDictionary(g => g.Key, g => g.Count());


            // Get action names from the `action` collection
            var actionNames = actionCollection.Find(new BsonDocument()).ToList()
                .ToDictionary(doc => doc["_id"].AsInt32, doc => doc["name"].AsString);

            if (actions.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu hành động!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Prepare data for the bar char
            double[] values = actions.Values.Select(v => (double)v).ToArray();
            string[] labels = actions.Keys.Select(id => actionNames.ContainsKey(id) ? actionNames[id] : $"Hành động {id}").ToArray();
            double[] positions = Enumerable.Range(0, values.Length).Select(i => (double)i).ToArray();

            // Clear the old plot
            formsPlot1.Plot.Clear();

            // Create and add bars
            var bars = new List<ScottPlot.Bar>();
            for (int i = 0; i < values.Length; i++)
            {
                bars.Add(new ScottPlot.Bar
                {
                    Position = positions[i],
                    Value = values[i],
                    FillColor = ScottPlot.Colors.Gray, // Correct for 5.x
                    Size = 0.6 // Correct for 5.x
                });
            }
            formsPlot1.Plot.Add.Bars(bars);

            // Set X-axis ticks
            formsPlot1.Plot.Axes.Bottom.SetTicks(positions, labels); // Correct for 5.x

            // Set Y-axis limits (start from 0)
            formsPlot1.Plot.Axes.SetLimitsY(0, values.Max() * 1.1); // Correct for 5.x

            // Add titles and labels
            formsPlot1.Plot.Title("Biểu đồ hành động");
            formsPlot1.Plot.YLabel("Số lần thực hiện");

            // Refresh the plot
            formsPlot1.Refresh();
        }

    }
}
