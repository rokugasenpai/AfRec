﻿/*--------------------------------------------------------------------------
* AfRec
* ver 1.0.3.0 (2015/03/25)
*
* Copyright © 2015 Rokugasenpai All Rights Reserved.
* licensed under Microsoft Public License(Ms-PL)
* https://github.com/rokugasenpai
* https://twitter.com/rokugasenpai
* http://nicorec.info
*--------------------------------------------------------------------------*/

﻿/*--------------------------------------------------------------------------
* 使用したライブラリのクレジット
*
* DynamicJson
* ver 1.2.0.0 (May. 21th, 2010)
*
* created and maintained by neuecc <ils@neue.cc>
* licensed under Microsoft Public License(Ms-PL)
* http://neue.cc/
* http://dynamicjson.codeplex.com/
*
* FFmpeg 32bit Shared
* ver N-69672-g078be09
*
* FFmpeg License
* FFmpeg is licensed under the GNU Lesser General Public License (LGPL) version 2.1 or later. However, FFmpeg incorporates several optional parts and optimizations that are covered by the GNU General Public License (GPL) version 2 or later. If those parts get used the GPL applies to all of FFmpeg.
* Read the license texts to learn how this affects programs built on top of FFmpeg or reusing FFmpeg. You may also wish to have a look at the GPL FAQ.
* Note that FFmpeg is not available under any other licensing terms, especially not proprietary/commercial ones, not even in exchange for payment.
*--------------------------------------------------------------------------*/

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Diagnostics;

using Codeplex.Data;

﻿/*--------------------------------------------------------------------------
* 処理概要
* ざっくりとした内容なので、必ずコードを見てください。
*
* ○ 処理開始(フォームの表示)
* ↓ 
* < FFmpegはあるか？ > → ○ エラー表示、処理終了
* ↓ 
* < Config.xmlはあるか？ > → [ Config.xmlの生成と               ]
* │                          [ デフォルトの保存先フォルダの作成 ]
* ↓
* I ユーザーがテキストボックスにURLを入力 I
* ↓
* I ユーザーが開始ボタンを押下            I
* ↓                                        ↓
* [ 放送タイトル、放送者名を取得するため  ] I ユーザーがキャンセルボタンを押下 I
* [ APIを叩いて、jsonをダウンロード       ] ↓
* ↓                                        [ 処理をキャンセル                 ]
* < 成功？ > → [一定時間後リトライ]
* ↓ 
* [ jsonより動画データであるtsファイルの  ]
* [ URLが記載されているm3u8ファイルを     ]
* [ ダウンロードするためのAPIキーを抽出   ]
* ↓ 
* ／ すべて終わるまでループ              ＼
* [ m3u8ファイルをダウンロード            ]
* ↓
* < 成功？ > → [一定時間後リトライ]
* ↓ 
* ＼                                     ／
* ↓
* [ tsファイルのURLを抽出                 ]
* ↓
* ／ すべて終わるまでループ              ＼
* [ tsファイルをダウンロード              ]
* ↓
* < 成功？ > → [一定時間後リトライ]
* ↓
* ＼                                     ／
* ↓
* [ tsファイルの数より                    ]
* [ 必要なディスク容量を算出              ]
* ↓
* < 必要なディスク容量？ > → ○ エラー表示、処理終了
* ↓
* ／ すべて終わるまでループ              ＼
* [ FFmpegを使い100ずつtsファイルを結合   ]
* ↓
* < 成功？ > → ○ エラー表示、処理終了
* ↓
* ＼                                     ／
* ↓
* < 結合ファイルが複数？ > → [ FFmpegを使いmp4変換 ]
* ↓                                              │
* [ FFmpegを使い最終的な結合・mp4変換     ]       │
* │←──────────────────────┘
* ↓
* < 成功？ > → ○ エラー表示、処理終了
* ↓
* ○ 成功表示、処理終了
*--------------------------------------------------------------------------*/

namespace AfRec
{

    public partial class FormMain : Form
    {
        private readonly String CONFIG_XML_PATH = Directory.GetParent(Application.ExecutablePath).FullName + Path.DirectorySeparatorChar + "Config.xml";
        private readonly String FFMPEG_EXE_PATH = Directory.GetParent(Application.ExecutablePath).FullName + Path.DirectorySeparatorChar + "ffmpeg.exe";
        private readonly String START_BUTTON_TEXT = "開始";
        private readonly String CANCEL_BUTTON_TEXT = "キャンセル";
        private readonly String NORMAL_MESSAGE_PREFIX = "【 正常 】";
        private readonly String ERROR_MESSAGE_PREFIX = "【エラー】";
        // ダウンロードエラー(WebClientのWebException)のリトライ。
        private readonly Int32 RETRY_TIMES = 1;
        private readonly Int32 RETRY_INTERVAL = 5000;

        // 処理の開始が可能かどうか。例えばFFmpegがなければtrue。
        private Boolean isUnstartableError = false;
        private String saveToPath = "";
        // ダウンロードしたTS形式の動画ファイルなどの一時的なファイルが置かれる。
        private String tempPath = "";
        // 入力されたURLから判別できる放送主のIDを表す。
        private String id = "";
        // 入力されたURLから判別できる放送のIDを表す。
        private String vno = "";
        // ダウンロードしたjsonファイルから判別できる放送タイトルを表す。
        private String title = "";
        // ダウンロードしたjsonファイルから判別できる放送主の名前を表す。
        private String nick = "";

        protected delegate void VoidCallback();
        protected delegate void SetStringCallback(String arg);

        protected void SetButtonText(String text)
        {
            try
            {
                if (InvokeRequired)
                {
                    SetStringCallback callback = new SetStringCallback(SetButtonText);
                    Invoke(callback, new object[] { text });
                    return;
                }
                button.Text = text;
            }
            catch (Exception)
            {

            }
        }

        protected void AppendTextTextBoxMessage(String text)
        {
            try
            {
                if (InvokeRequired)
                {
                    SetStringCallback callback = new SetStringCallback(AppendTextTextBoxMessage);
                    Invoke(callback, new object[] { text });
                    return;
                }
                textBoxMessage.AppendText(text);
            }
            catch (Exception)
            {

            }
        }

        protected void UpdateMessageText()
        {
            try
            {
                if (InvokeRequired)
                {
                    VoidCallback callback = new VoidCallback(UpdateMessageText);
                    Invoke(callback);
                    return;
                }
                textBoxMessage.Update();
            }
            catch (Exception)
            {

            }
        }

        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            Icon = Properties.Resources.afrec;
            button.Text = START_BUTTON_TEXT;

            if (!File.Exists(FFMPEG_EXE_PATH))
            {
                textBoxMessage.Text += ERROR_MESSAGE_PREFIX + "ffmepg.exeが見つかりませんでした。" + Environment.NewLine;
                this.isUnstartableError = true;
                return;
            }

            // 保存先情報を格納するxmlの処理
            this.saveToPath = Directory.GetParent(Application.ExecutablePath).FullName + Path.DirectorySeparatorChar + "Rec";
            XmlDocument xmlDoc = new XmlDocument();
            if (!File.Exists(CONFIG_XML_PATH))
            {
                xmlDoc.AppendChild(xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "yes"));
                XmlElement config = xmlDoc.CreateElement("config");
                XmlElement saveTo = xmlDoc.CreateElement("save_to");
                saveTo.InnerText = this.saveToPath;
                config.AppendChild(saveTo);
                xmlDoc.AppendChild(config);
                xmlDoc.Save(CONFIG_XML_PATH);
                textBoxMessage.Text += NORMAL_MESSAGE_PREFIX + "Config.xmlが無かったため新規作成しました。" + Environment.NewLine;
            }
            else
            {
                try
                {
                    xmlDoc.Load(CONFIG_XML_PATH);
                    if (xmlDoc.DocumentElement.GetElementsByTagName("save_to").Count == 0)
                    {
                        xmlDoc = new XmlDocument();
                        xmlDoc.AppendChild(xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "yes"));
                        XmlElement config = xmlDoc.CreateElement("config");
                        XmlElement saveTo = xmlDoc.CreateElement("save_to");
                        saveTo.InnerText = this.saveToPath;
                        config.AppendChild(saveTo);
                        xmlDoc.AppendChild(config);
                        xmlDoc.Save(CONFIG_XML_PATH);
                        textBoxMessage.Text += NORMAL_MESSAGE_PREFIX + "Config.xmlが壊れていたため新規作成しました。" + Environment.NewLine;
                    }
                    else
                    {
                        if (!Directory.Exists(xmlDoc.DocumentElement.GetElementsByTagName("save_to").Item(0).InnerText))
                        {
                            xmlDoc.DocumentElement.GetElementsByTagName("save_to").Item(0).InnerText = this.saveToPath;
                        }
                        else
                        {
                            this.saveToPath = xmlDoc.DocumentElement.GetElementsByTagName("save_to").Item(0).InnerText;
                        }
                    }
                }
                catch (Exception)
                {
                    textBoxMessage.Text += ERROR_MESSAGE_PREFIX + "Config.xmlの読み込みに失敗しました。" + Environment.NewLine;
                    this.isUnstartableError = true;
                    return;
                }

                if (!Directory.Exists(this.saveToPath))
                {
                    try
                    {
                        Directory.CreateDirectory(this.saveToPath);
                    }
                    catch (Exception)
                    {
                        textBoxMessage.Text += ERROR_MESSAGE_PREFIX + "保存先フォルダーの作成に失敗しました。" + Environment.NewLine;
                        this.isUnstartableError = true;
                        return;
                    }
                }

                textBoxSaveTo.Text = this.saveToPath;
            }
        }

        private void textBoxSaveTo_Enter(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxSaveTo.Text = folderBrowserDialog.SelectedPath;
                textBoxSaveTo.Update();
                XmlDocument xmlDoc = new XmlDocument();
                if (!File.Exists(CONFIG_XML_PATH))
                {
                    xmlDoc.AppendChild(xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "yes"));
                    XmlElement config = xmlDoc.CreateElement("config");
                    XmlElement saveTo = xmlDoc.CreateElement("save_to");
                    saveTo.InnerText = folderBrowserDialog.SelectedPath;
                    config.AppendChild(saveTo);
                    xmlDoc.AppendChild(config);
                    xmlDoc.Save(CONFIG_XML_PATH);
                    textBoxMessage.Text += NORMAL_MESSAGE_PREFIX + "Config.xmlが無かったため新規作成しました。" + Environment.NewLine;
                }
                else
                {
                    xmlDoc.DocumentElement.GetElementsByTagName("save_to").Item(0).InnerText = folderBrowserDialog.SelectedPath;
                    xmlDoc.Save(CONFIG_XML_PATH);
                }
            }
        }

        private void button_Click(object sender, EventArgs e)
        {
            if (button.Text == START_BUTTON_TEXT)
            {
                if (this.isUnstartableError)
                {
                    return;
                }
                Regex regex = new Regex(@"^\s*?http://afreecatv\.jp/(\d+?)/v/(\d+?)\D*$", RegexOptions.Compiled);
                Match match = regex.Match(textBoxUrl.Text);
                if (!match.Success)
                {
                    textBoxMessage.Text += ERROR_MESSAGE_PREFIX + "URLが正しいか確認して下さい。" + Environment.NewLine;
                }
                else
                {
                    this.id = match.Groups[1].Value;
                    this.vno = match.Groups[2].Value;
                    worker.RunWorkerAsync();
                }
            }
            else if (button.Text == CANCEL_BUTTON_TEXT)
            {
                worker.CancelAsync();
            }
            else
            {

            }
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            SetButtonText(CANCEL_BUTTON_TEXT);

            try
            {
                this.tempPath = this.saveToPath + Path.DirectorySeparatorChar + "Temp";
                if (Directory.Exists(this.tempPath))
                {
                    Directory.Delete(this.tempPath, true);
                }
                try
                {
                    Directory.CreateDirectory(this.tempPath);
                }
                catch (Exception)
                {
                    textBoxMessage.Text += ERROR_MESSAGE_PREFIX + "一時フォルダーの作成に失敗しました。" + Environment.NewLine;
                    this.isUnstartableError = true;
                    return;
                }

                String json = "";

                // APIよりjsonをダウンロード
                using (WebClient wc1 = new WebClient())
                {
                    Int32 cnt = 0;
                    String url = "http://api.afreecatv.jp/video/view_video.php";

                    // POSTのパラメータの生成
                    NameValueCollection postData = new NameValueCollection();
                    postData.Add("vno", this.vno);
                    postData.Add("rt", "json");
                    postData.Add("lc", "ja_JP");
                    postData.Add("bid", this.id);
                    postData.Add("pt", "view");
                    postData.Add("cptc", "HLS");

                    while (cnt <= RETRY_TIMES)
                    {
                        try
                        {
                            json = Encoding.UTF8.GetString(wc1.UploadValues(url, postData));
                            if (worker.CancellationPending)
                            {
                                AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "中止しました。" + Environment.NewLine);
                                return;
                            }
                            break;
                        }
                        catch (WebException webEx)
                        {
                            if (cnt == RETRY_TIMES)
                            {
                                AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + url + "のダウンロードで問題が発生しました。" + webEx.Message + Environment.NewLine);
                                return;
                            }
                            System.Threading.Thread.Sleep(RETRY_INTERVAL);
                            cnt++;
                        }
                    }

                    AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "動画情報" + new Uri(url).Segments.Last() + "をダウンロードしました。" + Environment.NewLine);
                    UpdateMessageText();
                }

                List<String> m3u8List = new List<string>();
                List<String> chatList = new List<string>();
                Dictionary<String, List<String>> tsDict = new Dictionary<String, List<String>>();

                // jsonよりm3u8ファイルのURLを抽出
                try
                {
                    dynamic obj = DynamicJson.Parse(json);
                    this.title = obj.channel.title;
                    this.nick = obj.channel.nick;

                    foreach (var elem in obj.channel.flist)
                    {
                        String file = elem.file;
                        String chat = elem.chat;
                        chatList.Add(chat);
                        String fileBase = file.Replace("/" + new Uri(file).Segments.Last(), "");
                        tsDict.Add(fileBase, new List<String>());
                        m3u8List.Add(fileBase + "/index_0_av.m3u8");
                    }

                }
                catch (Exception ex)
                {
                    AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + "動画URLの解析で問題が発生しました。" + ex.Message + Environment.NewLine);
                    return;
                }

                foreach (String chat in chatList)
                {

                    // 動画情報m3u8ファイルのダウンロード
                    using (WebClient wc = new WebClient())
                    {
                        Int32 cnt = 0;
                        while (cnt <= RETRY_TIMES)
                        {
                            try
                            {
                                wc.DownloadFile(chat, this.saveToPath + Path.DirectorySeparatorChar + SanitizeFileName(this.vno + " - " + this.nick + " - " + this.title + " - " + (chatList.IndexOf(chat) + 1).ToString() + ".xml"));
                                
                                if (worker.CancellationPending)
                                {
                                    AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "中止しました。" + Environment.NewLine);
                                    return;
                                }

                                break;
                            }
                            catch (WebException webEx)
                            {
                                if (cnt == RETRY_TIMES)
                                {
                                    AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + chat + "のダウンロードで問題が発生しました。" + webEx.Message + Environment.NewLine);
                                    // チャットログは落とせなくても続行とする。
                                }
                                System.Threading.Thread.Sleep(RETRY_INTERVAL);
                                cnt++;
                            }
                        }

                        AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + (chatList.IndexOf(chat) + 1).ToString() + " / " + chatList.Count + " チャットログをダウンロードしました。" + Environment.NewLine);
                        UpdateMessageText();
                    }
                }

                Int32 numTs = 0;
                Int32 numFinTs = 0;

                foreach (String m3u8 in m3u8List)
                {

                    using (WebClient wc = new WebClient())
                    {
                        Int32 cnt = 0;
                        String key = m3u8.Replace("/index_0_av.m3u8", "");
                        String res = "";

                        while (cnt <= RETRY_TIMES)
                        {
                            try
                            {
                                res = Encoding.UTF8.GetString(wc.DownloadData(m3u8));

                                if (worker.CancellationPending)
                                {
                                    AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "中止しました。" + Environment.NewLine);
                                    return;
                                }
                                break;
                            }
                            catch (WebException webEx)
                            {
                                if (cnt == RETRY_TIMES)
                                {
                                    AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + m3u8 + "のダウンロードで問題が発生しました。" + webEx.Message + Environment.NewLine);
                                    return;
                                }
                                System.Threading.Thread.Sleep(RETRY_INTERVAL);
                                cnt++;
                            }
                        }

                        AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "動画情報" + new Uri(key).Segments.Last() + "をダウンロードしました。" + Environment.NewLine);
                        UpdateMessageText();

                        // m3u8ファイルよりtsファイルのURLを抽出
                        Regex regex = new Regex(@"(http://.+)", RegexOptions.Compiled);
                        MatchCollection matches = regex.Matches(res);
                        foreach (Match match in matches)
                        {
                            if (match.Success)
                            {
                                tsDict[key].Add(match.Groups[1].Value);
                                numTs++;
                            }
                        }
                    }

                }

                // tsファイルの数より必要なディスク容量を算出
                String driveLetter = Application.ExecutablePath.Substring(0, 1);
                DriveInfo drive = new DriveInfo(driveLetter);
                Int64 mb = (Int64)Math.Pow(1024, 2);
                Int64 gb = (Int64)Math.Pow(1024, 3);
                Int64 shortage = (numTs * (3 * mb) * 2) - drive.AvailableFreeSpace;
                if (shortage > 0)
                {
                    Double shortageGb = Math.Round((Double)shortage / gb, 2);
                    AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + driveLetter + "ドライブの空き容量が" + shortageGb.ToString() + "GB不足しています。" + Environment.NewLine);
                    return;
                }

                AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "動画ファイルのダウンロードを開始します。" + Environment.NewLine);
                UpdateMessageText();
                foreach (KeyValuePair<String, List<String>> kv in tsDict)
                {
                    foreach(String ts in kv.Value)
                    {

                        // tsファイルのダウンロード
                        using (WebClient wc = new WebClient())
                        {
                            Int32 cnt = 0;
                            String zeroAdded = "0000" + (numFinTs + 1).ToString();
                            String fileName = zeroAdded.Substring(zeroAdded.Length - 4, 4) + ".ts";

                            while (cnt <= RETRY_TIMES)
                            {
                                try
                                {
                                    wc.DownloadFile(ts, this.tempPath + Path.DirectorySeparatorChar + fileName);
                                    numFinTs++;

                                    AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + numFinTs.ToString() + " / " + numTs.ToString() + " 動画ファイルをダウンロードしました。" + Environment.NewLine);
                                    UpdateMessageText();
                                    
                                    if (worker.CancellationPending)
                                    {
                                        AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "中止しました。" + Environment.NewLine);
                                        return;
                                    }
                                    break;
                                }
                                catch (WebException webEx)
                                {
                                    // 末尾のTSファイルが無い場合があるので、404だったらbreakする。
                                    if (webEx.Status == WebExceptionStatus.ProtocolError)
                                    {
                                        if (((HttpWebResponse)webEx.Response).StatusCode == HttpStatusCode.NotFound)
                                        {
                                            break;
                                        }
                                    }
                                    if (cnt == RETRY_TIMES)
                                    {
                                        AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + ts + "のダウンロードで問題が発生しました。" + webEx.Message + Environment.NewLine);
                                        return;
                                    }
                                    System.Threading.Thread.Sleep(RETRY_INTERVAL);
                                    cnt++;
                                }
                            }
                        }

                    }
                }

                AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "動画ファイルの結合を開始します。" + Environment.NewLine);
                UpdateMessageText();

                try
                {
                    FileInfo[] files = new DirectoryInfo(this.tempPath).GetFiles();

                    // ファイル名昇順にソート
                    Array.Sort<FileInfo>(files, delegate(FileInfo x, FileInfo y)
                    {
                        return x.Name.CompareTo(y.Name);
                    });

                    List<String> tsFiles = new List<String>();
                    List<String> concatFiles = new List<String>();
                    Int32 numConcat = 100;
                    Int32 cntConcat = 0;

                    foreach (FileInfo file in files)
                    {

                        if (Regex.Match(file.Name, @"\d+\.ts").Success)
                        {
                            tsFiles.Add(file.FullName);
                            cntConcat++;
                            if (tsFiles.Count == numConcat)
                            {

                                // tsファイルを100ずつ結合
                                // 一度に全部結合させない理由は、FFmpegのパラメータが長すぎると、
                                // Windowsの仕様で問題が起こるため
                                // https://support.microsoft.com/ja-jp/kb/2823587/ja
                                using (Process ps = new Process())
                                {
                                    String fileName = (cntConcat - numConcat + 1).ToString() + "_" + cntConcat.ToString() + ".ts";
                                    String filePath = this.tempPath + Path.DirectorySeparatorChar + fileName;

                                    ps.StartInfo.FileName = FFMPEG_EXE_PATH;
                                    // concat:を使うと無劣化結合できる
                                    ps.StartInfo.Arguments = "-y -i \"concat:" + String.Join("|", tsFiles.ToArray()) + "\" -c copy \"" + filePath + "\"";
                                    ps.StartInfo.CreateNoWindow = true;
                                    ps.StartInfo.UseShellExecute = false;
                                    ps.Start();
                                    System.Threading.Thread.Sleep(1000);

                                    while (true)
                                    {
                                        if (ps.HasExited)
                                        {
                                            foreach (String ts in tsFiles)
                                            {
                                                File.Delete(ts);
                                            }
                                            tsFiles = new List<String>();

                                            if (ps.ExitCode != 0)
                                            {
                                                AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + cntConcat.ToString() + " / " + numFinTs.ToString() + " 動画ファイルを結合でエラーが発生しました。" + Environment.NewLine);
                                            }
                                            else
                                            {
                                                cntConcat += tsFiles.Count;
                                                concatFiles.Add(filePath);
                                                AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + cntConcat.ToString() + " / " + numFinTs.ToString() + " 動画ファイルを結合しました。" + Environment.NewLine);
                                            }
                                            UpdateMessageText();

                                            break;
                                        }

                                        if (worker.CancellationPending)
                                        {
                                            ps.Kill();
                                            AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "中止しました。" + Environment.NewLine);
                                            return;
                                        }

                                        System.Threading.Thread.Sleep(100);
                                    }
                                    
                                }
                            }
                        }
                    }

                    // 100ずつ処理した後の端数のtsファイルを結合
                    if (tsFiles.Count < numConcat)
                    {

                        using (Process ps = new Process())
                        {
                            String fileName = (cntConcat - (cntConcat % numConcat) + 1).ToString() + "_" + cntConcat.ToString() + ".ts";
                            String filePath = this.tempPath + Path.DirectorySeparatorChar + fileName;

                            ps.StartInfo.FileName = FFMPEG_EXE_PATH;
                            ps.StartInfo.Arguments = "-y -i \"concat:" + String.Join("|", tsFiles.ToArray()) + "\" -c copy \"" + filePath + "\"";
                            ps.StartInfo.CreateNoWindow = true;
                            ps.StartInfo.UseShellExecute = false;
                            ps.Start();
                            System.Threading.Thread.Sleep(1000);

                            while (true)
                            {
                                if (ps.HasExited)
                                {
                                    foreach (String ts in tsFiles)
                                    {
                                        File.Delete(ts);
                                    }
                                    tsFiles = new List<String>();

                                    if (ps.ExitCode != 0)
                                    {
                                        AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + cntConcat.ToString() + " / " + numFinTs.ToString() + " 動画ファイルを結合でエラーが発生しました。" + Environment.NewLine);
                                    }
                                    else
                                    {
                                        cntConcat += tsFiles.Count;
                                        concatFiles.Add(filePath);
                                        AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + cntConcat.ToString() + " / " + numFinTs.ToString() + " 動画ファイルを結合しました。" + Environment.NewLine);
                                    }
                                    UpdateMessageText();

                                    break;
                                }

                                if (worker.CancellationPending)
                                {
                                    ps.Kill();
                                    AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "中止しました。" + Environment.NewLine);
                                    return;
                                }

                                System.Threading.Thread.Sleep(100);
                            }
                        }
                    }

                    // 出力されるファイル名の形式
                    // 放送番号 - 放送者名 - 放送タイトル.mp4
                    String completeFileName = SanitizeFileName(this.vno + " - " + this.nick + " - " + this.title + ".mp4");

                    using (Process ps = new Process())
                    {
                        String filePath = this.saveToPath + Path.DirectorySeparatorChar + completeFileName;

                        ps.StartInfo.FileName = FFMPEG_EXE_PATH;
                        if (concatFiles.Count > 1)
                        {
                            ps.StartInfo.Arguments = "-y -i \"concat:" + String.Join("|", concatFiles.ToArray()) + "\" -c copy -bsf:a aac_adtstoasc \"" + filePath + "\"";
                        }
                        else
                        {
                            ps.StartInfo.Arguments = "-y -i \"" + concatFiles[0] + "\" -c copy -bsf:a aac_adtstoasc \"" + filePath + "\"";
                        }
                        ps.StartInfo.CreateNoWindow = true;
                        ps.StartInfo.UseShellExecute = false;
                        ps.Start();
                        System.Threading.Thread.Sleep(1000);

                        while (true)
                        {
                            if (ps.HasExited)
                            {
                                foreach (String concat in concatFiles)
                                {
                                    File.Delete(concat);
                                }
                                concatFiles = new List<String>();

                                if (ps.ExitCode != 0)
                                {
                                    AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + "全動画ファイルの結合・MP4変換でエラーが発生しました。" + Environment.NewLine);
                                }
                                else
                                {
                                    AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "完了しました。ファイル名：" + completeFileName + Environment.NewLine);
                                }
                                UpdateMessageText();

                                break;
                            }

                            if (worker.CancellationPending)
                            {
                                ps.Kill();
                                AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "中止しました。" + Environment.NewLine);
                                return;
                            }

                            System.Threading.Thread.Sleep(100);
                        }
                    }

                }
                catch (Exception ex)
                {
                    AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + "動画ファイルの処理で問題が発生しました。" + ex.Message + Environment.NewLine);
                    return;
                }
            }
            catch (Exception ex)
            {
                AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + "問題が発生しました。" + ex.Message + Environment.NewLine);
                return;
            }
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 処理完了(正常＆エラー)時の後処理
            textBoxUrl.Text = "";
            button.Text = START_BUTTON_TEXT;
            try
            {
                Directory.Delete(this.tempPath, true);
            }
            catch (Exception)
            {

            }
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (worker.IsBusy)
            {
                if (MessageBox.Show("実行中ですが終了しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    worker.CancelAsync();
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

        // 放送タイトルをファイル名に使用するため、ファイル名として使えない文字のサニタイズを行う
        private String SanitizeFileName(String fileName)
        {
            fileName = fileName.Replace("/", "／").Replace("\\", "￥").Replace("?", "？").Replace("*", "＊")
                .Replace(":", "：").Replace("\"", "”").Replace("<", "＜").Replace(">", "＞");

            if (fileName.Last() == '.')
            {
                fileName = fileName.Substring(0, fileName.Length - 1) + "．";
            }

            if (fileName.Last() == ' ')
            {
                fileName = fileName.Substring(0, fileName.Length - 1) + "　";
            }

            return fileName;
        }
    }
}
