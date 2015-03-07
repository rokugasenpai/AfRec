﻿/*--------------------------------------------------------------------------
* AfRec
* ver 1.0.0.0 (2015/03/07)
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
        private readonly Int32 RETRY_TIMES = 1;
        private readonly Int32 RETRY_INTERVAL = 5000;

        private Boolean isError = false;
        private String saveToPath = "";
        private String tempPath = "";
        private String id = "";
        private String vno = "";
        private String title = "";
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
                this.isError = true;
                return;
            }

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
                    this.isError = true;
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
                        this.isError = true;
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
                if (this.isError)
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
                    this.isError = true;
                    return;
                }

                String json = "";

                using (WebClient wc1 = new WebClient())
                {
                    Int32 cnt = 0;
                    String url = "http://api.afreecatv.jp/video/view_video.php";

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

                String driveLetter = Application.ExecutablePath.Substring(0, 1);
                DriveInfo drive = new DriveInfo(driveLetter);
                Int64 mb = (Int64)Math.Pow(1024, 2);
                Int64 gb = (Int64)Math.Pow(1024, 3);
                Int64 shortage = (numTs * (2 * mb) * 2) - drive.AvailableFreeSpace;
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

                                    AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + numFinTs.ToString() + " / " + numTs.ToString() + " 動画ファイルのダウンロード中です。" + Environment.NewLine);
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
                                            numFinTs++;
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

                AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "動画ファイルの結合、MP4変換を開始します。" + Environment.NewLine);
                UpdateMessageText();
                String completeFileName = SanitizeFileName(this.vno + " - " + this.nick + " - " + this.title + ".mp4");

                // コマンドラインでffmpegのconcatを利用する際、パラメーターが長すぎるとエラーが出るため
                // 一定数で連結させ、その複数あるかもしれない連結TSファイルを、1つのMP4ファイルに結合する。
                // ffmpegのconcatは元ファイルがTS形式でないと結合できない。

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
                    Int32 numConcat = 500;
                    Int32 cntConcat = 0;

                    foreach (FileInfo file in files)
                    {

                        if (Regex.Match(file.Name, @"\d+\.ts").Success)
                        {
                            tsFiles.Add(file.FullName);
                            if (tsFiles.Count == numConcat)
                            {

                                using (Process ps = new Process())
                                {
                                    String fileName = (cntConcat - numConcat + 1).ToString() + "_" + cntConcat.ToString() + ".ts";
                                    String filePath = this.tempPath + Path.DirectorySeparatorChar + fileName;

                                    ps.StartInfo.FileName = FFMPEG_EXE_PATH;
                                    ps.StartInfo.Arguments = "-y -i \"concat:" + String.Join("|", tsFiles.ToArray()) + "\" -c copy \"" + filePath + "\"";
                                    ps.StartInfo.CreateNoWindow = true;
                                    ps.StartInfo.UseShellExecute = false;
                                    ps.Start();
                                    System.Threading.Thread.Sleep(5000);

                                    while (true)
                                    {
                                        if (ps.HasExited)
                                        {
                                            cntConcat += tsFiles.Count;
                                            concatFiles.Add(filePath);

                                            foreach (String ts in tsFiles)
                                            {
                                                File.Delete(ts);
                                            }
                                            tsFiles = new List<String>();

                                            AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + cntConcat.ToString() + " / " + numFinTs.ToString() + " 動画ファイルの結合・MP4変換中です。" + Environment.NewLine);
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
                            System.Threading.Thread.Sleep(5000);

                            while (true)
                            {
                                if (ps.HasExited)
                                {
                                    cntConcat += tsFiles.Count;
                                    concatFiles.Add(filePath);

                                    foreach (String ts in tsFiles)
                                    {
                                        File.Delete(ts);
                                    }
                                    tsFiles = new List<String>();

                                    AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + cntConcat.ToString() + " / " + numFinTs.ToString() + " 動画ファイルの結合・MP4変換中です。" + Environment.NewLine);
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

                    if (concatFiles.Count == 1)
                    {
                        File.Move(concatFiles[0], completeFileName);
                    }
                    else
                    {

                        using (Process ps = new Process())
                        {
                            String filePath = this.saveToPath + Path.DirectorySeparatorChar + completeFileName;

                            ps.StartInfo.FileName = FFMPEG_EXE_PATH;
                            ps.StartInfo.Arguments = "-y -i \"concat:" + String.Join("|", concatFiles.ToArray()) + "\" -c copy -bsf:a aac_adtstoasc \"" + filePath + "\"";
                            ps.StartInfo.CreateNoWindow = true;
                            ps.StartInfo.UseShellExecute = false;
                            ps.Start();
                            System.Threading.Thread.Sleep(5000);

                            while (true)
                            {
                                if (ps.HasExited)
                                {
                                    AppendTextTextBoxMessage(NORMAL_MESSAGE_PREFIX + "完了しました。ファイル名：" + completeFileName + Environment.NewLine);
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
                catch (Exception ex)
                {
                    AppendTextTextBoxMessage(ERROR_MESSAGE_PREFIX + "動画ファイルの結合、MP4変換で問題が発生しました。" + ex.Message + Environment.NewLine);
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
