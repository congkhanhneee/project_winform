using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MongoDB.Bson;
using MongoDB.Driver;

namespace WindowsFormsApp2
{
    public partial class Form3 : Form
    {
        private readonly IMongoCollection<BsonDocument> fileTypeCollection;
        private readonly IMongoCollection<BsonDocument> identityCollection;
        private bool isDatabaseConnected = false; // Cờ kiểm tra kết nối database

        public Form3()
        {
            InitializeComponent();
            try
            {
                var client = new MongoClient("mongodb://localhost:27017");
                var database = client.GetDatabase("FileManagement");

                fileTypeCollection = database.GetCollection<BsonDocument>("file_type");
                identityCollection = database.GetCollection<BsonDocument>("identity");

                isDatabaseConnected = true; // Đánh dấu kết nối thành công
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể kết nối CSDL: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            if (!isDatabaseConnected) return; // Nếu không có kết nối, thoát khỏi hàm

            dt_File.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dt_File.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dt_File.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            LoadFileTypes();
            LoadExtensions();
        }

        private void LoadFileTypes()
        {
            if (!isDatabaseConnected) return;

            try
            {
                var fileTypes = fileTypeCollection.Find(new BsonDocument()).ToList();
                cbTypeFile.Items.Clear();
                foreach (var type in fileTypes)
                {
                    cbTypeFile.Items.Add(type["name"].AsString);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải danh sách loại tệp: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadExtensions()
        {
            if (!isDatabaseConnected) return;

            try
            {
                var extensions = identityCollection.Find(new BsonDocument()).ToList();
                cbExtensionFile.Items.Clear();
                foreach (var ext in extensions)
                {
                    cbExtensionFile.Items.Add(ext["_id"].AsString);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải danh sách phần mở rộng: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadFilteredFiles()
        {
            if (!isDatabaseConnected) return;

            string directoryPath = @"D:\"; // Thư mục cần kiểm tra
            if (!Directory.Exists(directoryPath))
            {
                MessageBox.Show("Thư mục không tồn tại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                var files = new List<string>();
                try
                {
                    files = GetAllFilesSafe(directoryPath).ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi quét thư mục: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                lb_Sum_Directory.Text = $"Tổng số tệp trong thư mục: {files.Count}";

                var selectedType = cbTypeFile.SelectedItem?.ToString();
                if (!string.IsNullOrEmpty(selectedType))
                {
                    var fileTypeDoc = fileTypeCollection.Find(Builders<BsonDocument>.Filter.Eq("name", selectedType)).FirstOrDefault();
                    if (fileTypeDoc == null)
                    {
                        MessageBox.Show($"Không tìm thấy loại tệp '{selectedType}' trong database!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    int fileTypeId = fileTypeDoc["_id"].AsInt32;
                    var extensionsForType = identityCollection
                        .Find(Builders<BsonDocument>.Filter.Eq("file_type_id", fileTypeId))
                        .ToList()
                        .Select(doc => doc["_id"].AsString)
                        .ToList();

                    files = files.Where(f => extensionsForType.Contains(Path.GetExtension(f).ToLower())).ToList();
                }

                var selectedExtension = cbExtensionFile.SelectedItem?.ToString();
                if (!string.IsNullOrEmpty(selectedExtension))
                {
                    files = files.Where(f => Path.GetExtension(f).Equals(selectedExtension, StringComparison.OrdinalIgnoreCase)).ToList();
                }

                lb_Sum_File.Text = $"Tổng số tệp: {files.Count}";
                dt_File.DataSource = files.Select(f => new
                {
                    Name = Path.GetFileName(f),
                    Path = f,
                    Extension = Path.GetExtension(f)
                }).ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi quét thư mục: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Quét tất cả file trong thư mục (bao gồm thư mục con) và bỏ qua lỗi quyền truy cập
        /// </summary>
        private IEnumerable<string> GetAllFilesSafe(string rootPath)
        {
            var files = new List<string>();

            try
            {
                files.AddRange(Directory.GetFiles(rootPath, "*.*", SearchOption.TopDirectoryOnly));
            }
            catch (UnauthorizedAccessException) { } // Bỏ qua lỗi không có quyền
            catch (IOException) { } // Bỏ qua lỗi khi thư mục không hợp lệ

            try
            {
                foreach (var directory in Directory.GetDirectories(rootPath))
                {
                    files.AddRange(GetAllFilesSafe(directory)); // Đệ quy quét thư mục con
                }
            }
            catch (UnauthorizedAccessException) { }
            catch (IOException) { }

            return files;
        }



        private void btn_Click_Directory_Click(object sender, EventArgs e)
        {
            LoadFilteredFiles();
        }
    }
}
