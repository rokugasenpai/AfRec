using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("---------- Start ----------");
            using (Process ps = new Process())
            {
                ps.StartInfo.FileName = "ffmpeg";
                ps.StartInfo.Arguments = @"-y -i C:\souonojisan.wav -unko libmp3lame -ab 192k C:\souonojisan.mp3";
                ps.StartInfo.CreateNoWindow = true;
                ps.StartInfo.UseShellExecute = false;
                ps.StartInfo.RedirectStandardError = true;
                ps.Start();
                
                ps.WaitForExit();
                if (ps.ExitCode != 0)
                {
                    String stderr = ps.StandardError.ReadToEnd();
                    String errmsg = "";
                    foreach (String line in stderr.Split(new String[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (!line.StartsWith("ffmpeg version") && !line.StartsWith("  "))
                        {
                            errmsg += line + Environment.NewLine;
                        }
                    }
                    Console.WriteLine("FFmpeg error.");
                    Console.WriteLine(errmsg);
                }
            }
            Console.WriteLine("---------- Finish ----------");
            Console.WriteLine("Press Enter key.");
            Console.ReadLine();
        }

        public static void CopyDirectory(string sourceDirName, string destDirName)
        {
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
                File.SetAttributes(destDirName,
                    File.GetAttributes(sourceDirName));
            }
            if (destDirName[destDirName.Length - 1] !=
                    Path.DirectorySeparatorChar)
            {
                destDirName = destDirName + Path.DirectorySeparatorChar;
            }
            string[] files = Directory.GetFiles(sourceDirName);
            foreach (string file in files)
            {
                File.Copy(file, destDirName + Path.GetFileName(file), true);
            }
            string[] dirs = Directory.GetDirectories(sourceDirName);
            foreach (string dir in dirs)
            {
                CopyDirectory(dir, destDirName + Path.GetFileName(dir));
            }
        }
    }
}
