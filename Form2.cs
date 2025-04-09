using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using MongoDB.Bson;
using MongoDB.Driver;

namespace WindowsFormsApp2
{
    public partial class Form2 : Form
    {
        private IMongoDatabase database;
        private IMongoCollection<BsonDocument> detailFileCollection;
        private IMongoCollection<BsonDocument> fileCollection;
        private IMongoCollection<BsonDocument> actionCollection;
        private IMongoCollection<BsonDocument> userCollection;

        private DateTime startTime;
        private DateTime endTime;
        private bool isDatabaseConnected = false; // Cờ kiểm tra kết nối DB

        public Form2(DateTime start, DateTime end)
        {
            InitializeComponent();
            startTime = start;
            endTime = end;

            // Kiểm tra kết nối MongoDB
            try
            {
                var client = new MongoClient("mongodb://localhost:27017");
                database = client.GetDatabase("FileManagement");

                // Kiểm tra xem có thể truy cập được collection không
                detailFileCollection = database.GetCollection<BsonDocument>("detail_file");
                fileCollection = database.GetCollection<BsonDocument>("file");
                actionCollection = database.GetCollection<BsonDocument>("action");
                userCollection = database.GetCollection<BsonDocument>("user");

                isDatabaseConnected = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể kết nối CSDL: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if (!isDatabaseConnected) return; // Nếu không kết nối được DB thì thoát
            LoadStatistics();
            LoadFileList();
        }

        private void LoadStatistics()
        {
            if (!isDatabaseConnected) return;

            try
            {
                var filter = Builders<BsonDocument>.Filter.Gte("datetime", startTime.ToString("yyyy-MM-dd HH:mm:ss")) &
                             Builders<BsonDocument>.Filter.Lte("datetime", endTime.ToString("yyyy-MM-dd HH:mm:ss"));

                var detailFiles = detailFileCollection.Find(filter).ToList();
                int totalActions = detailFiles.Count;

                var actionDocs = actionCollection.Find(new BsonDocument()).ToList();
                var actionDict = actionDocs.ToDictionary(a => a["_id"].ToString(), a => a["name"].ToString());

                int createdFiles = detailFiles.Count(d => actionDict.TryGetValue(d["id_action"].ToString(), out string action) && action == "Tạo");
                int editedFiles = detailFiles.Count(d => actionDict.TryGetValue(d["id_action"].ToString(), out string action) && action == "Sửa");
                int deletedFiles = detailFiles.Count(d => actionDict.TryGetValue(d["id_action"].ToString(), out string action) && action == "Xóa");
                int renamedFiles = detailFiles.Count(d => actionDict.TryGetValue(d["id_action"].ToString(), out string action) && action.StartsWith("Đổi tên"));
                int openedFiles = detailFiles.Count(d => actionDict.TryGetValue(d["id_action"].ToString(), out string action) && action == "Mở");

                lblTotalActions.Text = $"Tổng số hành động: {totalActions}";
                lblCreatedFiles.Text = $"Số file được tạo: {createdFiles}";
                lblOpenedFiles.Text = $"Số file được mở: {openedFiles}";
                lblEditedFiles.Text = $"Số file được sửa: {editedFiles}";
                lblDeletedFiles.Text = $"Số file bị xóa: {deletedFiles}";
                lblRenamedFiles.Text = $"Số file bị đổi tên: {renamedFiles}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải thống kê: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadFileList()
        {
            if (!isDatabaseConnected) return;

            try
            {
                var filter = Builders<BsonDocument>.Filter.Gte("datetime", startTime.ToString("yyyy-MM-dd HH:mm:ss")) &
                             Builders<BsonDocument>.Filter.Lte("datetime", endTime.ToString("yyyy-MM-dd HH:mm:ss"));

                var fileIds = detailFileCollection.Find(filter)
                                  .Project(Builders<BsonDocument>.Projection.Include("id_file"))
                                  .ToList()
                                  .Select(d => d["id_file"].AsObjectId)
                                  .Distinct()
                                  .ToList();

                if (!fileIds.Any()) return;

                cboFiles.Items.Clear();

                foreach (var fileId in fileIds)
                {
                    try
                    {
                        var fileDoc = fileCollection.Find(Builders<BsonDocument>.Filter.Eq("_id", fileId)).FirstOrDefault();
                        if (fileDoc != null && fileDoc.Contains("name"))
                        {
                            cboFiles.Items.Add(fileDoc["name"].ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Lỗi khi xử lý fileId {fileId}: {ex.Message}");
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải danh sách tệp: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDetails_File_Click(object sender, EventArgs e)
        {
            if (!isDatabaseConnected) return;

            if (cboFiles.SelectedItem == null)
            {
                MessageBox.Show("Vui lòng chọn một tệp!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string selectedFile = cboFiles.SelectedItem.ToString();
            LoadFileHistory(selectedFile);
        }

        private void LoadFileHistory(string fileName)
        {
            if (!isDatabaseConnected) return;
            try
            {
                var fileFilter = Builders<BsonDocument>.Filter.Eq("name", fileName);
                var fileDoc = fileCollection.Find(fileFilter).FirstOrDefault();

                if (fileDoc == null)
                {
                    MessageBox.Show("Không tìm thấy tệp!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var detailFilter = Builders<BsonDocument>.Filter.Eq("id_file", fileDoc["_id"].AsObjectId);
                var fileHistory = detailFileCollection.Find(detailFilter)
                                  .Sort(Builders<BsonDocument>.Sort.Descending("datetime"))
                                  .ToList();

                DataTable dt = new DataTable();
                dt.Columns.Add("DATE_ACTION");
                dt.Columns.Add("CURRENT_NAME");
                dt.Columns.Add("OLD_NAME");
                dt.Columns.Add("OLD_PATH");
                dt.Columns.Add("NEW_PATH");
                dt.Columns.Add("NAME_ACTION");
                dt.Columns.Add("NAME_USER");

                foreach (var detail in fileHistory)
                {
                    // Lấy id_action (chuyển về số)
                    string actionName = "Không xác định";
                    if (detail.Contains("id_action"))
                    {
                        int idAction;
                        if (int.TryParse(detail["id_action"].ToString(), out idAction)) // Chuyển chuỗi thành số
                        {
                            var actionFilter = Builders<BsonDocument>.Filter.Eq("_id", idAction);
                            var actionDoc = actionCollection.Find(actionFilter).FirstOrDefault();
                            if (actionDoc != null)
                            {
                                actionName = actionDoc["name"].ToString();
                            }
                        }
                    }

                    // Lấy id_user (chuyển về số)
                    string userName = "Không xác định";

                    if (detail.Contains("id_user"))
                    {
                        string idUserString = detail["id_user"].ToString(); // Lấy id_user dưới dạng chuỗi
                        try
                        {
                            ObjectId idUser = ObjectId.Parse(idUserString); // Chuyển thành ObjectId

                            var userFilter = Builders<BsonDocument>.Filter.Eq("_id", idUser);
                            var userDoc = userCollection.Find(userFilter).FirstOrDefault();

                            if (userDoc != null && userDoc.Contains("name"))
                            {
                                userName = userDoc["name"].AsString;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Lỗi chuyển ObjectId: {ex.Message}");
                        }
                    }

                    dt.Rows.Add(
                        detail["datetime"].ToString(),
                        fileDoc["name"].ToString(),
                        detail.Contains("oldname") ? detail["oldname"].ToString() : "Không xác định",
                        detail.Contains("old_path") ? detail["old_path"].ToString() : "Không xác định",
                        detail.Contains("new_path") ? detail["new_path"].ToString() : "Không xác định",
                        actionName,
                        userName
                    );
                }

                dgvFileHistory.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải lịch sử tệp: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
