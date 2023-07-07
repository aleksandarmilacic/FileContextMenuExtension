using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace FileContextMenuExtension
{
    public class CopyFilePathContextMenu
    {
        private const string CopyAsCommaSeparated = "Copy as Comma-Separated";
        private const string CopyAsNewLineSeparated = "Copy as New Line Separated";

        // Import the necessary Win32 APIs
        [DllImport("user32.dll")]
        public static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("user32.dll")]
        public static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        public static extern bool CloseClipboard();

        [DllImport("user32.dll")]
        public static extern bool EmptyClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetClipboardData(uint uFormat, IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll")]
        public static extern UIntPtr GlobalSize(IntPtr hMem);

        public static void Main(string[] args)
        {
            // Retrieve the selected file paths
            string[] selectedPaths = Environment.GetCommandLineArgs();

            // Copy the paths based on the chosen option
            if (args.Length > 0)
            {
                string verb = args[0];
                
                if (verb == CopyAsCommaSeparated)
                {
                    string pathsText = string.Join(",", selectedPaths[1..]);
                    CopyToClipboard(pathsText);
                }
                else if (verb == CopyAsNewLineSeparated)
                {
                    string pathsText = string.Join(Environment.NewLine, selectedPaths[1..]);
                    CopyToClipboard(pathsText);
                }
            }
        }

        private static void CopyToClipboard(string text)
        {
            try
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    EmptyClipboard();

                    // Allocate memory for the string
                    byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text + "\0");
                    UIntPtr dataSize = new UIntPtr((uint)textBytes.Length);
                    IntPtr hGlobal = GlobalAlloc(0x2000, dataSize);

                    if (hGlobal != IntPtr.Zero)
                    {
                        IntPtr lockedMemory = GlobalLock(hGlobal);

                        if (lockedMemory != IntPtr.Zero)
                        {
                            // Copy the text to the allocated memory
                            Marshal.Copy(textBytes, 0, lockedMemory, textBytes.Length);

                            GlobalUnlock(lockedMemory);
                            SetClipboardData(13, hGlobal); // 13 is CF_UNICODETEXT

                            CloseClipboard();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to copy text to clipboard: {ex.Message}");
            }
        }
    }
}
