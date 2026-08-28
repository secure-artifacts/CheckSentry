#nullable enable
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace CheckSentry.Launcher;

internal static class Program
{
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int MessageBox(IntPtr handle, string text, string caption, uint type);

    private static int Main(string[] args)
    {
        string logPath = string.Empty;
        try
        {
            var forwarded = args.Where(a => !string.Equals(a, "--portable", StringComparison.OrdinalIgnoreCase)).ToArray();
            var dataDirectory = Path.GetFullPath(AppContext.BaseDirectory);
            Directory.CreateDirectory(dataDirectory);
            MigrateLegacyData(dataDirectory);
            var logDirectory = Path.Combine(dataDirectory, "Logs");
            Directory.CreateDirectory(logDirectory);
            logPath = Path.Combine(logDirectory, "CheckSentry-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".log");

            var assembly = typeof(Program).Assembly;
            var runtimeDirectory = Path.Combine(dataDirectory, ".CheckSentryRuntime", assembly.ManifestModule.ModuleVersionId.ToString("N"));
            ExtractPayload(assembly, runtimeDirectory);
            var scriptPath = Path.Combine(runtimeDirectory, "Start-ComplianceCheck.ps1");
            var listPath = Path.Combine(dataDirectory, "list.xlsx");

            var startInfo = new ProcessStartInfo
            {
                FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe"),
                UseShellExecute = false,
                WorkingDirectory = runtimeDirectory,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };
            foreach (var argument in new[] { "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptPath, "-ListPath", listPath })
                startInfo.ArgumentList.Add(argument);
            foreach (var argument in forwarded)
                startInfo.ArgumentList.Add(argument);

            using var logWriter = new StreamWriter(logPath, append: false, new UTF8Encoding(false)) { AutoFlush = true };
            using var process = new Process { StartInfo = startInfo, EnableRaisingEvents = true };
            process.OutputDataReceived += (_, eventArgs) => WriteProcessLine(eventArgs.Data, logWriter, false);
            process.ErrorDataReceived += (_, eventArgs) => WriteProcessLine(eventArgs.Data, logWriter, true);
            if (!process.Start()) throw new InvalidOperationException("无法启动 Windows PowerShell。 ");
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            process.WaitForExit();
            process.WaitForExit(1000);
            if (process.ExitCode != 0)
            {
                var message = "CheckSentry 启动失败。\n\n错误日志：\n" + logPath;
                MessageBox(IntPtr.Zero, message, "CheckSentry", 0x10);
            }
            return process.ExitCode;
        }
        catch (Exception exception)
        {
            var message = "CheckSentry 启动失败：" + exception.Message;
            try
            {
                if (!string.IsNullOrWhiteSpace(logPath)) File.AppendAllText(logPath, message + Environment.NewLine, new UTF8Encoding(false));
                MessageBox(IntPtr.Zero, message + (string.IsNullOrWhiteSpace(logPath) ? string.Empty : "\n\n错误日志：\n" + logPath), "CheckSentry", 0x10);
            }
            catch { Console.Error.WriteLine(message); }
            return 1;
        }
    }

    private static void ExtractPayload(Assembly assembly, string runtimeDirectory)
    {
        Directory.CreateDirectory(runtimeDirectory);
        var root = Path.GetFullPath(runtimeDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        foreach (var resourceName in assembly.GetManifestResourceNames().Where(name => name.StartsWith("payload/", StringComparison.Ordinal)))
        {
            var relative = resourceName["payload/".Length..].Replace('/', Path.DirectorySeparatorChar);
            var destination = Path.GetFullPath(Path.Combine(runtimeDirectory, relative));
            if (!destination.StartsWith(root, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("嵌入资源路径无效。 ");
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            using var source = assembly.GetManifestResourceStream(resourceName) ?? throw new InvalidDataException("无法读取嵌入资源：" + resourceName);
            var temporary = destination + ".tmp-" + Guid.NewGuid().ToString("N");
            using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)) source.CopyTo(output);
            File.Move(temporary, destination, true);
        }
        if (!File.Exists(Path.Combine(runtimeDirectory, "Start-ComplianceCheck.ps1"))) throw new InvalidDataException("自包含资源不完整。 ");
    }

    private static void MigrateLegacyData(string dataDirectory)
    {
        var legacyDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CheckSentry");
        if (string.Equals(Path.GetFullPath(legacyDirectory).TrimEnd(Path.DirectorySeparatorChar), dataDirectory.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase)) return;
        if (!Directory.Exists(legacyDirectory)) return;

        foreach (var name in new[] { "list.xlsx", "CheckSentry.settings.json", "CheckSentry.cloud.json" })
        {
            var source = Path.Combine(legacyDirectory, name);
            var destination = Path.Combine(dataDirectory, name);
            if (!File.Exists(source)) continue;
            if (!File.Exists(destination))
            {
                File.Move(source, destination);
            }
            else
            {
                var recoveredName = Path.GetFileNameWithoutExtension(name) + "-recovered-from-C-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + Path.GetExtension(name);
                File.Move(source, Path.Combine(dataDirectory, recoveredName));
            }
        }

        var legacyLogs = Path.Combine(legacyDirectory, "Logs");
        if (Directory.Exists(legacyLogs))
        {
            var targetLogs = Path.Combine(dataDirectory, "Logs");
            Directory.CreateDirectory(targetLogs);
            foreach (var source in Directory.GetFiles(legacyLogs, "*.log"))
            {
                var destination = Path.Combine(targetLogs, Path.GetFileName(source));
                if (File.Exists(destination)) destination = Path.Combine(targetLogs, Path.GetFileNameWithoutExtension(source) + "-migrated-" + Guid.NewGuid().ToString("N") + ".log");
                File.Move(source, destination);
            }
            if (!Directory.EnumerateFileSystemEntries(legacyLogs).Any()) Directory.Delete(legacyLogs);
        }

        var legacyRuntime = Path.Combine(legacyDirectory, "Runtime");
        if (Directory.Exists(legacyRuntime)) Directory.Delete(legacyRuntime, recursive: true);
        if (!Directory.EnumerateFileSystemEntries(legacyDirectory).Any()) Directory.Delete(legacyDirectory);
    }

    private static void WriteProcessLine(string? line, TextWriter logWriter, bool error)
    {
        if (line is null) return;
        lock (logWriter)
        {
            logWriter.WriteLine(line);
            if (error) Console.Error.WriteLine(line); else Console.WriteLine(line);
        }
    }
}
