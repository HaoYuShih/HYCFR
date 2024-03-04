using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace HYCFR
{
    public class CFunc
    {
        #region 民國西元轉換string To string
        /// <summary>
        /// 此Func用為在Parse成DateTime前先使用字串做民國和西元的年份轉換，避免閏年時會造成轉換錯誤!
        /// 輸入 : 2012/5/12 => 101/5/12 , 101/5/12 => 2012/5/12 ，務必使用"/"或"-"!!!!
        /// </summary>
        /// <param name="time">輸入民國或西元格式</param>
        /// <param name="IsBC">true : 西元，false : 民國</param>
        /// <returns>(string)西元或民國</returns>
        public string BCswitchTW(string time, bool IsBC = true)
        {
            string result = string.Empty;
            char splichar = '/';
            try
            {
                if (string.IsNullOrEmpty(time))
                {
                    return result;
                }

                //日期分隔符號判斷
                if (time.Substring(3, 1) == "/" || time.Substring(4, 1) == "/")
                {
                    splichar = '/';
                }
                else if (time.Substring(3, 1) == "-" || time.Substring(4, 1) == "-")
                {
                    splichar = '-';
                }
                else
                {
                    return result;
                }

                if (IsBC) //西元To民國
                {
                    int firIndex = time.IndexOf(splichar);
                    string str_year = time.Substring(0, firIndex);
                    int int_year;
                    if (!Int32.TryParse(str_year, out int_year))
                    {
                        throw new Exception("西元To民國時，轉換數字發生錯誤!");
                    }
                    int tw_year = int_year - 1911;
                    result = tw_year.ToString() + time.Substring(firIndex);
                }
                else //民國To西元
                {
                    int firIndex = time.IndexOf(splichar);
                    string str_year = time.Substring(0, firIndex);
                    int int_year;
                    if (!Int32.TryParse(str_year, out int_year))
                    {
                        throw new Exception("民國To西元時，轉換數字發生錯誤!");
                    }
                    int bc_year = int_year + 1911;
                    result = bc_year.ToString() + time.Substring(firIndex);
                }
                return result;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        #endregion

        #region 副檔名MimeType轉換
        /// <summary>
        /// 輸入副檔名或Mimetype進行轉換
        /// </summary>
        /// <param name="value">副檔名或Mimetype</param>
        /// <returns>(string)副檔名或Mimetype</returns>
        public string ExtSwitchMimeType(string value)
        {
            string result = string.Empty;
            try
            {
                #region 附檔名MimeType對照表
                Dictionary<string, string> dic = new Dictionary<string, string>()
                {
                    {".doc","application/msword" },
                    {".docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
                    {".gif","image/gif" },
                    {".svg","image/svg+xml" },
                    {".jpg","image/jpeg" },
                    {".jpeg","image/jpeg" },
                    {".png","image/png" },
                    {".pdf","application/pdf" },
                    {".ppt","application/vnd.ms-powerpoint" },
                    {".pptx","application/vnd.openxmlformats-officedocument.presentationml.presentation" },
                    {".rar","application/vnd.rar" },
                    {".txt","text/plain" },
                    {".xls","application/vnd.ms-excel" },
                    {".xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                    {".zip","application/zip" },
                    {".7z","application/x-7z-compressed" },
                };
                #endregion

                #region 轉換判斷
                if (value.Substring(0, 1) == ".")
                {
                    if (dic.ContainsKey(value))
                    {
                        result = dic[value];
                    }
                }
                else
                {
                    if (dic.ContainsValue(value))
                    {
                        result = dic.FirstOrDefault(x => x.Value == value).Key ?? string.Empty;
                    }
                }
                #endregion

                return result;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        #endregion

        #region 鹽產生器
        public string CreateSalt()
        {
            var buffer = new byte[4];
            var rng = new RNGCryptoServiceProvider(); //亂數產生器
            rng.GetBytes(buffer);
            rng.Dispose();
            return Convert.ToBase64String(buffer);
        }
        #endregion

        #region SHA256雜湊
        /// <summary>
        /// 輸入字串做SHA256雜湊
        /// </summary>
        /// <param name="value">(string)</param>
        /// <returns>(string)</returns>
        public string Hash_SHA256(string value)
        {
            HashAlgorithm ha = SHA256.Create();
            byte[] BytesData = Encoding.Default.GetBytes(value);
            byte[] BytesHash = ha.ComputeHash(BytesData);
            string result = BitConverter.ToString(BytesHash);
            return result;
        }
        #endregion

        #region 判斷Int是否為指定區間的數字
        /// <summary>
        /// 判斷Int介於指定區間內，並指定轉換型態string/Int
        /// </summary>
        /// <typeparam name="T">(string/Int)</typeparam>
        /// <param name="value">數字</param>
        /// <param name="startnum">大於等於此參數</param>
        /// <param name="endnum">小於等於此參數</param>
        /// <param name="result">(string/Int)value</param>
        /// <returns>(bool)是否驗證成功</returns>
        public bool CheckInt<T, Q>(T value, int startnum, int endnum, out Q result)
        {
            try
            {
                int target;
                result = default(Q);
                if (value == null)
                {
                    return false;
                }
                switch (typeof(T).Name)
                {
                    case nameof(String):
                        if (int.TryParse(value.ToString().Trim(), out target))
                        {
                            if (target >= startnum && target <= endnum)
                            {
                                break;
                            }
                        }
                        return false;
                    case nameof(Int32):
                        if (int.TryParse(value.ToString().Trim(), out target))
                        {
                            if (target >= startnum && target <= endnum)
                            {
                                break;
                            }
                        }
                        return false;
                    default:
                        return false;
                }

                if (typeof(Q) == typeof(string))
                {
                    result = (Q)(object)(target.ToString());
                }
                if (typeof(Q) == typeof(int))
                {
                    result = (Q)(object)target;
                }

                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }
        #endregion

        #region 檢查字串字數在指定範圍內
        /// <summary>
        /// 檢查字串字數在指定長度範圍內
        /// </summary>
        /// <param name="target">字串</param>
        /// <param name="startnum">大於等於此參數</param>
        /// <param name="endnum">小於等於此參數</param>
        /// <param name="result">回傳target</param>
        /// <returns>(bool)是否驗證成功</returns>
        public bool CheckString(string target, int startnum, int endnum, out string result)
        {
            result = string.Empty;
            if (target != null)
            {
                if (target.Trim().Length >= startnum && target.Trim().Length <= endnum)
                {
                    result = target.Trim();
                    return true;
                }
            }
            return false;
        }
        #endregion

        #region 新ID產生器-可自訂格式
        /// <summary>
        /// 產生下一號ID 2 10
        /// 範例: formatExpress => "WL!{yyyyMM}IA%{4}" => WL20240115IA0001
        /// 代號說明:
        /// !{日期格式} => !{yyyyMMdd} 
        /// %{序號位數} => SN%{4}T => SN0001T 
        /// </summary>
        /// <param name="formatExpress">格式樣板</param>
        /// <param name="LatestID">最新ID</param>
        /// <returns>回傳下一號ID</returns>
        public string CreateNextID(string formatExpress, string LatestID)
        {
            try
            {
                StringBuilder strBuil = new StringBuilder();

                for (int i = 0; i < formatExpress.Length; i++)
                {
                    #region 處理日期
                    if (formatExpress[i] == '!')
                    {
                        int datelength = formatExpress.IndexOf('}', i);
                        if (datelength == -1)
                        {
                            throw new Exception("缺少'}'結束符號");
                        }
                        string dateformat = formatExpress.Substring(i + 2, datelength - i - 2);
                        string date = DateTime.Now.ToString(dateformat);
                        strBuil.Append(date);
                        i = datelength;
                        continue;
                    }
                    #endregion

                    #region 處理序號
                    if (formatExpress[i] == '%')
                    {
                        int intlength = formatExpress.IndexOf('}', i);
                        if (intlength == -1)
                        {
                            throw new Exception("缺少'}'結束符號");
                        }
                        string intcount = formatExpress.Substring(i + 2, intlength - i - 2);
                        if (!Int32.TryParse(intcount, out int count))
                        {
                            throw new Exception("%{}參數不是數字");
                            //i = intlength;
                            //continue;
                        }

                        if (string.IsNullOrEmpty(LatestID))
                        {
                            string intStr = "1".PadLeft(count, '0');
                            strBuil.Append(intStr);
                            i = intlength;
                            continue;
                        }
                        else
                        {
                            if (strBuil.Length + count > LatestID.Length)
                            {
                                throw new Exception("LatestID序號參數長度與formatExpress參數不符");
                            }
                            string nowIndex = LatestID.Substring(strBuil.Length, count);
                            if (!Int32.TryParse(nowIndex, out int nowindex_int))
                            {
                                throw new Exception("LatestID序號參數不是數字");
                            }
                            nowindex_int++;
                            strBuil.Append(nowindex_int.ToString().PadLeft(count, '0'));
                            i = intlength;
                            continue;
                        }
                    }
                    #endregion

                    strBuil.Append(formatExpress[i]);
                }

                return strBuil.ToString();
            }
            catch (Exception)
            {
                throw;
            }
        }
        #endregion

        #region 建好指定路徑資料夾
        /// <summary>
        /// 建好指定路徑資料夾
        /// </summary>
        /// <param name="path">完整路徑</param>
        /// <returns>(bool)是否有建資料夾</returns>
        /// <exception cref="Exception"></exception>
        public bool CreateDirectory(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        #endregion

        #region 清除檔案名稱中包含多重附檔名
        public string RemoveMultiExten(string filename)
        {
            string namewithoutExten = Path.GetFileNameWithoutExtension(filename);
            if (namewithoutExten.IndexOf('.') >= 0)
            {
                namewithoutExten = namewithoutExten.Substring(0, namewithoutExten.IndexOf('.'));
            }
            return namewithoutExten + Path.GetExtension(filename).ToLower();
        }
        #endregion

        #region 檢查副檔名
        public bool checkExtension(List<string> allowExten, string filename)
        {
            return allowExten.Contains(Path.GetExtension(filename));
        }
        #endregion

        #region 檔案類別
        public class FileData
        {
            public MemoryStream stream { get; set; }
            public string FileName { get; set; }
        }
        #endregion

        #region 檔案上傳
        /// <summary>
        /// 可批量上傳檔案至同一個路徑
        /// </summary>
        /// <param name="files">檔案</param>                                                     
        /// <param name="path">路徑</param>
        /// <param name="filenamelist">實際存檔檔名</param>
        /// <param name="IsRenew">是否先清空當前目錄下所有檔案</param>
        /// <param name="UseIndex">是否使用序號改寫檔名，避免重複檔名問題</param>
        /// <returns>成功上傳回傳空字串，反之回傳錯誤資訊</returns>
        public string UploadFile(List<FileData> files, string path, out List<string> filenamelist, bool IsRenew = false, bool UseIndex = false)
        {
            #region 變數
            filenamelist = new List<string>();
            string reStr = string.Empty;
            #endregion

            #region 更新資料夾檔案
            try
            {
                #region 是否清除目前資料夾的所有檔案
                if (IsRenew)
                {
                    Directory.Delete(path, true);
                    Directory.CreateDirectory(path);
                }
                #endregion

                int index = 1;
                foreach (var item in files)
                {
                    #region 檔名處理
                    string clearname = string.Empty;

                    if (UseIndex)
                    {
                        clearname = string.Concat(index.ToString(), Path.GetExtension(item.FileName).ToLower());
                    }
                    else
                    {
                        clearname = RemoveMultiExten(item.FileName); //清除多重附檔名
                    }
                    #endregion

                    #region 檢查檔名有沒有重複
#if true
                    //重新命名
                    int count = 1;
                    while (File.Exists(Path.Combine(path, clearname)))
                    {
                        clearname = string.Concat(
                            Path.GetFileNameWithoutExtension(clearname) + "_" + count.ToString(),
                            Path.GetExtension(clearname));
                        count++;
                    }
#endif

#if false
                    //有重複直接刪掉
                    if (File.Exists(Path.Combine(TargetDirFullPath, clearname)))
                    {
                        File.Delete(Path.Combine(TargetDirFullPath, clearname));
                    } 
#endif
                    #endregion

                    #region 上傳檔案
                    string finalpath = Path.Combine(path, clearname);
                    using (FileStream fs = File.Create(finalpath))
                    {
                        item.stream.Seek(0, SeekOrigin.Begin);
                        item.stream.CopyTo(fs);
                    }
                    #endregion

                    filenamelist.Add(clearname);
                    index++;
                }

                return reStr;
            }
            catch (Exception ex)
            {
                reStr = "檔案上傳過程出錯!";
                return reStr;
            }
            #endregion
        }
        /// <summary>
        /// 可批量上傳檔案至同一個路徑
        /// </summary>
        /// <param name="files">檔案</param>
        /// <param name="path">路徑</param>
        /// <param name="filenamelist">實際存檔檔名</param>
        /// <param name="IsRenew">是否先清空當前目錄下所有檔案</param>
        /// <param name="UseIndex">是否使用序號改寫檔名，避免重複檔名問題</param>
        /// <returns>成功上傳回傳空字串，反之回傳錯誤資訊</returns>
        public string UploadFile(List<HttpPostedFileBase> files, string path, out List<string> filenamelist, bool IsRenew = false, bool UseIndex = false)
        {
            #region 變數
            filenamelist = new List<string>();
            string reStr = string.Empty;
            #endregion

            #region 更新資料夾檔案
            try
            {
                #region 是否清除目前資料夾的所有檔案
                if (IsRenew)
                {
                    Directory.Delete(path, true);
                    Directory.CreateDirectory(path);
                }
                #endregion

                int index = 1;
                foreach (var item in files)
                {
                    #region 檔名處理
                    string clearname = string.Empty;

                    if (UseIndex)
                    {
                        clearname = string.Concat(index.ToString(), Path.GetExtension(item.FileName).ToLower());
                    }
                    else
                    {
                        clearname = RemoveMultiExten(item.FileName); //清除多重附檔名
                    }
                    #endregion

                    #region 檢查檔名有沒有重複
#if true
                    //重新命名
                    int count = 1;
                    while (File.Exists(Path.Combine(path, clearname)))
                    {
                        clearname = string.Concat(
                            Path.GetFileNameWithoutExtension(clearname) + "_" + count.ToString(),
                            Path.GetExtension(clearname));
                        count++;
                    }
#endif

#if false
                    //有重複直接刪掉
                    if (File.Exists(Path.Combine(TargetDirFullPath, clearname)))
                    {
                        File.Delete(Path.Combine(TargetDirFullPath, clearname));
                    } 
#endif
                    #endregion

                    #region 上傳檔案
                    string finalpath = Path.Combine(path, clearname);
                    item.SaveAs(finalpath);
                    #endregion

                    filenamelist.Add(clearname);
                    index++;
                }

                return reStr;
            }
            catch (Exception ex)
            {
                reStr = "檔案上傳過程出錯!";
                return reStr;
            }
            #endregion
        }
        #endregion

        #region 檔案下載
        /// <summary>
        /// 單一檔案下載
        /// </summary>
        /// <param name="path">資料夾路徑</param>
        /// <param name="filename">檔案名稱</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public FileData DownloadFile(string path, string filename)
        {
            #region 變數
            FileData fs = new FileData();
            #endregion

            try
            {
                #region 抓檔案
                if (File.Exists(Path.Combine(path, filename)))
                {
                    fs.stream = new MemoryStream(File.ReadAllBytes(Path.Combine(path, filename)));
                    fs.FileName = filename;
                    return fs;
                }
                else
                {
                    return fs;
                }
                #endregion
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        #endregion

        #region 抓取資料夾內所有檔名
        /// <summary>
        /// 抓取資料夾內所有檔名
        /// </summary>
        /// <param name="path">資料夾路徑</param>
        /// <returns>檔名陣列</returns>
        /// <exception cref="Exception"></exception>
        public List<string> GetFileNamesFromDir(string path)
        {
            #region 變數
            List<string> result = new List<string>();
            #endregion

            try
            {
                #region 抓檔名
                if (Directory.Exists(path))
                {
                    foreach (var item in Directory.GetFiles(path))
                    {
                        result.Add(Path.GetFileName(item));
                    }
                }
                #endregion

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        #endregion

        #region Excel匯出
        /// <summary>
        /// 從DataGrid搜尋結果匯出Excel資料
        /// layout範例:
        /// {"狀態", "State"}
        /// </summary>
        /// <param name="layout">Excel的行標題與資料的Key對應表</param>
        /// <param name="datagridResult">DataDrid查詢資料結果</param>
        /// <param name="title">頁簽的標題</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public MemoryStream MakeExcel(Dictionary<string, string> layout, string datagridResult, string title)
        {
            try
            {
                JObject data = (JObject)(JsonConvert.DeserializeObject(datagridResult));
                JArray ja = data["rows"] != null ? (JArray)data["rows"] : null;

                var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add(title); //excel的標籤頁面

                #region 標題
                int ColumnIndex = 1;
                foreach (var Dic in layout)
                {
                    worksheet.Cells[1, ColumnIndex].Value = Dic.Key;
                    ColumnIndex++;
                }
                #endregion

                #region 內容
                if (ja != null)
                {
                    int RowIndex = 2;
                    foreach (var item in ja)
                    {
                        for (int i = 1; i < (layout.Count + 1); i++)
                        {
                            worksheet.Cells[RowIndex, i].Value = item[layout[worksheet.Cells[1, i].Value.ToString()]]?.ToString();
                        }
                        RowIndex++;
                    }
                }
                #endregion

                worksheet.Cells.AutoFitColumns();
                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Seek(0, SeekOrigin.Begin);
                return stream;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        #endregion
    }
}