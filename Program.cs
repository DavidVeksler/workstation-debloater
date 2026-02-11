using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using Microsoft.Win32;

// ── PerformanceFixer — Windows Console Utility ──────────────────────────────
// Automates performance fixes for Dell Precision 7680 dev workstations.
// Must be run as Administrator.

const string RestoreScriptPath = @"C:\Users\DVeksler\PerformanceFixer\restore.ps1";
var restoreLines = new List<string>
{
    "# PerformanceFixer Restore Script",
    $"# Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
    "# Run as Administrator to undo changes made by PerformanceFixer.",
    ""
};

if (!IsAdministrator())
{
    WriteColor("ERROR: This tool must be run as Administrator.", ConsoleColor.Red);
    WriteColor("Right-click your terminal and select 'Run as administrator', then retry.", ConsoleColor.Yellow);
    return 1;
}

WriteColor("╔══════════════════════════════════════════════════╗", ConsoleColor.Cyan);
WriteColor("║        PerformanceFixer  v1.0                   ║", ConsoleColor.Cyan);
WriteColor("║        Dell Precision 7680 Optimizer            ║", ConsoleColor.Cyan);
WriteColor("╚══════════════════════════════════════════════════╝", ConsoleColor.Cyan);
Console.WriteLine();

while (true)
{
    Console.WriteLine();
    WriteColor("────────────── MENU ──────────────", ConsoleColor.Cyan);
    Console.WriteLine("[0] Run ALL fixes (with confirmation per category)");
    Console.WriteLine("[1] Show system snapshot (before/after)");
    Console.WriteLine("[2] Kill orphaned dotnet.exe & VBCSCompiler processes");
    Console.WriteLine("[3] Switch power plan to High Performance");
    Console.WriteLine("[4] Stop & disable Dell bloatware services");
    Console.WriteLine("[5] Stop & disable redundant remote access services");
    Console.WriteLine("[6] Disable unnecessary scheduled tasks");
    Console.WriteLine("[7] Kill resource-heavy non-essential processes");
    Console.WriteLine("[8] Disable unnecessary startup entries");
    Console.WriteLine("[9] Add Windows Defender exclusions for dev paths");
    Console.WriteLine("[Q] Quit");
    Console.Write("\nSelect option: ");

    var key = Console.ReadLine()?.Trim().ToUpperInvariant();
    Console.WriteLine();

    switch (key)
    {
        case "0": await RunAllFixes(restoreLines); break;
        case "1": ShowSystemSnapshot(); break;
        case "2": KillOrphanedDotnetProcesses(restoreLines); break;
        case "3": await SwitchPowerPlan(restoreLines); break;
        case "4": StopDellBloatwareServices(restoreLines); break;
        case "5": StopRemoteAccessServices(restoreLines); break;
        case "6": DisableScheduledTasks(restoreLines); break;
        case "7": KillResourceHeavyProcesses(restoreLines); break;
        case "8": DisableStartupEntries(restoreLines); break;
        case "9": await AddDefenderExclusions(restoreLines); break;
        case "Q": goto done;
        default:
            WriteColor("Invalid option.", ConsoleColor.Yellow);
            break;
    }
}

done:
WriteRestoreScript(restoreLines);
WriteColor("Goodbye! Restore script saved to: " + RestoreScriptPath, ConsoleColor.Green);
return 0;

// ═══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════════

static bool IsAdministrator()
{
    using var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
    var principal = new System.Security.Principal.WindowsPrincipal(identity);
    return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
}

static void WriteColor(string msg, ConsoleColor color)
{
    var prev = Console.ForegroundColor;
    Console.ForegroundColor = color;
    Console.WriteLine(msg);
    Console.ForegroundColor = prev;
}

static void WriteInline(string msg, ConsoleColor color)
{
    var prev = Console.ForegroundColor;
    Console.ForegroundColor = color;
    Console.Write(msg);
    Console.ForegroundColor = prev;
}

static bool Confirm(string prompt)
{
    WriteInline($"  {prompt} [Y/n] ", ConsoleColor.Yellow);
    var answer = Console.ReadLine()?.Trim().ToUpperInvariant();
    return answer is "" or "Y" or "YES";
}

static string FormatBytes(long bytes)
{
    if (bytes < 1024) return $"{bytes} B";
    if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
    if (bytes < 1024L * 1024 * 1024) return $"{bytes / (1024.0 * 1024):F1} MB";
    return $"{bytes / (1024.0 * 1024 * 1024):F2} GB";
}

static async Task<(string output, int exitCode)> RunProcess(string fileName, string arguments)
{
    var psi = new ProcessStartInfo
    {
        FileName = fileName,
        Arguments = arguments,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false,
        CreateNoWindow = true
    };
    using var proc = Process.Start(psi)!;
    // Read both streams concurrently to avoid deadlock when a pipe buffer fills
    var stdoutTask = proc.StandardOutput.ReadToEndAsync();
    var stderrTask = proc.StandardError.ReadToEndAsync();
    await proc.WaitForExitAsync();
    string output = await stdoutTask;
    await stderrTask;
    return (output, proc.ExitCode);
}

static int? GetParentProcessId(Process p)
{
    try
    {
        using var query = new ManagementObjectSearcher(
            $"SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {p.Id}");
        foreach (var item in query.Get())
            return Convert.ToInt32(item["ParentProcessId"]);
    }
    catch { }
    return null;
}

static void WriteRestoreScript(List<string> lines)
{
    if (lines.Count <= 4) return; // only header, no actual restore commands
    try
    {
        File.WriteAllLines(RestoreScriptPath, lines);
    }
    catch (Exception ex)
    {
        WriteColor($"Warning: Could not write restore script: {ex.Message}", ConsoleColor.Yellow);
    }
}

static void SectionHeader(string title)
{
    WriteColor($"── {title} ──", ConsoleColor.Cyan);
}

// ═══════════════════════════════════════════════════════════════════════════════
// [0] RUN ALL FIXES
// ═══════════════════════════════════════════════════════════════════════════════

static async Task RunAllFixes(List<string> restore)
{
    SectionHeader("RUN ALL FIXES");
    WriteColor("This will walk through each fix category with confirmation.\n", ConsoleColor.White);

    ShowSystemSnapshot();
    Console.WriteLine();

    KillOrphanedDotnetProcesses(restore);
    Console.WriteLine();
    await SwitchPowerPlan(restore);
    Console.WriteLine();
    StopDellBloatwareServices(restore);
    Console.WriteLine();
    StopRemoteAccessServices(restore);
    Console.WriteLine();
    DisableScheduledTasks(restore);
    Console.WriteLine();
    KillResourceHeavyProcesses(restore);
    Console.WriteLine();
    DisableStartupEntries(restore);
    Console.WriteLine();
    await AddDefenderExclusions(restore);
    Console.WriteLine();

    WriteColor("All fix categories processed. Taking final snapshot...\n", ConsoleColor.Green);
    ShowSystemSnapshot();
}

// ═══════════════════════════════════════════════════════════════════════════════
// [1] SYSTEM SNAPSHOT
// ═══════════════════════════════════════════════════════════════════════════════

static void ShowSystemSnapshot()
{
    SectionHeader("SYSTEM SNAPSHOT");

    var processes = Process.GetProcesses();
    long totalRamUsed = 0;
    foreach (var p in processes)
    {
        try { totalRamUsed += p.WorkingSet64; } catch { }
    }

    Console.WriteLine($"  Total processes: {processes.Length}");
    Console.WriteLine($"  Total RAM used by processes: {FormatBytes(totalRamUsed)}");

    // Top 10 by RAM
    var topRam = processes
        .Select(p =>
        {
            try { return (Name: p.ProcessName, PID: p.Id, RAM: p.WorkingSet64); }
            catch { return (Name: p.ProcessName, PID: p.Id, RAM: 0L); }
        })
        .OrderByDescending(x => x.RAM)
        .Take(10)
        .ToList();

    Console.WriteLine("\n  Top 10 by RAM:");
    Console.WriteLine($"  {"Process",-30} {"PID",8} {"RAM",12}");
    foreach (var (name, pid, ram) in topRam)
        Console.WriteLine($"  {name,-30} {pid,8} {FormatBytes(ram),12}");

    // Top 10 by CPU (sample-based — take two snapshots 500ms apart)
    Console.Write("\n  Sampling CPU usage (1 second)...");
    var cpuStart = new Dictionary<int, TimeSpan>();
    foreach (var p in processes)
    {
        try { cpuStart[p.Id] = p.TotalProcessorTime; } catch { }
    }
    Thread.Sleep(1000);
    var cpuUsages = new List<(string Name, int PID, double CpuPct)>();
    int coreCount = Environment.ProcessorCount;
    foreach (var p in processes)
    {
        try
        {
            if (cpuStart.TryGetValue(p.Id, out var start))
            {
                var elapsed = p.TotalProcessorTime - start;
                double pct = elapsed.TotalMilliseconds / 1000.0 / coreCount * 100.0;
                cpuUsages.Add((p.ProcessName, p.Id, pct));
            }
        }
        catch { }
    }
    Console.WriteLine(" done.");

    var topCpu = cpuUsages.OrderByDescending(x => x.CpuPct).Take(10).ToList();
    Console.WriteLine("\n  Top 10 by CPU:");
    Console.WriteLine($"  {"Process",-30} {"PID",8} {"CPU %",8}");
    foreach (var (name, pid, cpu) in topCpu)
        Console.WriteLine($"  {name,-30} {pid,8} {cpu,7:F1}%");

    // Power plan
    try
    {
        var psi = new ProcessStartInfo("powercfg", "/getactivescheme")
        {
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var proc = Process.Start(psi)!;
        string line = proc.StandardOutput.ReadToEnd().Trim();
        proc.WaitForExit();
        Console.WriteLine($"\n  Active power plan: {line}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"\n  Power plan: (could not read: {ex.Message})");
    }

    // Running services
    var services = ServiceController.GetServices();
    int running = services.Count(s => s.Status == ServiceControllerStatus.Running);
    Console.WriteLine($"  Running services: {running} / {services.Length}");

    foreach (var p in processes)
        try { p.Dispose(); } catch { }
}

// ═══════════════════════════════════════════════════════════════════════════════
// [2] KILL ORPHANED DOTNET & VBCSCOMPILER
// ═══════════════════════════════════════════════════════════════════════════════

static void KillOrphanedDotnetProcesses(List<string> restore)
{
    SectionHeader("KILL ORPHANED dotnet.exe & VBCSCompiler");

    int killed = 0;
    long freed = 0;

    // Find devenv PIDs
    var devenvPids = new HashSet<int>();
    foreach (var p in Process.GetProcessesByName("devenv"))
    {
        devenvPids.Add(p.Id);
        p.Dispose();
    }

    // Kill dotnet.exe not parented by devenv
    foreach (var p in Process.GetProcessesByName("dotnet"))
    {
        try
        {
            int? parentId = GetParentProcessId(p);
            if (parentId.HasValue && devenvPids.Contains(parentId.Value))
            {
                WriteColor($"  Skipping dotnet.exe PID {p.Id} (child of devenv)", ConsoleColor.DarkGray);
                continue;
            }

            long mem = p.WorkingSet64;
            if (Confirm($"Kill dotnet.exe PID {p.Id} ({FormatBytes(mem)})?"))
            {
                p.Kill(entireProcessTree: true);
                p.WaitForExit(5000);
                killed++;
                freed += mem;
                WriteColor($"  Killed dotnet.exe PID {p.Id} — freed {FormatBytes(mem)}", ConsoleColor.Green);
            }
        }
        catch (Exception ex)
        {
            WriteColor($"  Error killing dotnet.exe PID {p.Id}: {ex.Message}", ConsoleColor.Red);
        }
        finally { p.Dispose(); }
    }

    // Kill VBCSCompiler
    foreach (var p in Process.GetProcessesByName("VBCSCompiler"))
    {
        try
        {
            long mem = p.WorkingSet64;
            if (Confirm($"Kill VBCSCompiler.exe PID {p.Id} ({FormatBytes(mem)})?"))
            {
                p.Kill();
                p.WaitForExit(5000);
                killed++;
                freed += mem;
                WriteColor($"  Killed VBCSCompiler.exe PID {p.Id} — freed {FormatBytes(mem)}", ConsoleColor.Green);
            }
        }
        catch (Exception ex)
        {
            WriteColor($"  Error killing VBCSCompiler PID {p.Id}: {ex.Message}", ConsoleColor.Red);
        }
        finally { p.Dispose(); }
    }

    WriteColor($"\n  Summary: killed {killed} process(es), freed ~{FormatBytes(freed)}", ConsoleColor.Green);
}

// ═══════════════════════════════════════════════════════════════════════════════
// [3] SWITCH POWER PLAN
// ═══════════════════════════════════════════════════════════════════════════════

static async Task SwitchPowerPlan(List<string> restore)
{
    SectionHeader("SWITCH POWER PLAN TO HIGH PERFORMANCE");

    const string highPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c";

    // Record current plan for restore
    var (currentPlan, _) = await RunProcess("powercfg", "/getactivescheme");
    currentPlan = currentPlan.Trim();
    string? currentGuid = null;
    // Extract GUID from output like "Power Scheme GUID: 381b4222-f694-41f0-9685-ff5bb260df2e  (Balanced)"
    var guidMatch = System.Text.RegularExpressions.Regex.Match(currentPlan, @"([0-9a-f\-]{36})", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    if (guidMatch.Success)
        currentGuid = guidMatch.Groups[1].Value;

    Console.WriteLine($"  Current plan: {currentPlan}");

    if (currentGuid?.Equals(highPerfGuid, StringComparison.OrdinalIgnoreCase) == true)
    {
        WriteColor("  Already on High Performance — no change needed.", ConsoleColor.Green);
        return;
    }

    if (!Confirm("Switch to High Performance plan?"))
        return;

    // Check if High Performance plan exists
    var (listOutput, _) = await RunProcess("powercfg", "/list");
    bool planExists = listOutput.Contains(highPerfGuid, StringComparison.OrdinalIgnoreCase);

    if (!planExists)
    {
        WriteColor("  High Performance plan not found. Creating it...", ConsoleColor.Yellow);
        var (dupOutput, dupCode) = await RunProcess("powercfg", $"/duplicatescheme {highPerfGuid}");
        if (dupCode != 0)
        {
            WriteColor($"  Failed to create High Performance plan: {dupOutput}", ConsoleColor.Red);
            return;
        }
        WriteColor("  Created High Performance plan.", ConsoleColor.Green);
    }

    var (_, exitCode) = await RunProcess("powercfg", $"/setactive {highPerfGuid}");
    if (exitCode == 0)
    {
        WriteColor("  Power plan switched to High Performance.", ConsoleColor.Green);
        if (currentGuid != null)
        {
            restore.Add($"# Restore original power plan");
            restore.Add($"powercfg /setactive {currentGuid}");
            restore.Add("");
        }
    }
    else
    {
        WriteColor("  Failed to switch power plan.", ConsoleColor.Red);
    }

    // Verify
    var (verify, _) = await RunProcess("powercfg", "/getactivescheme");
    Console.WriteLine($"  Verified: {verify.Trim()}");
}

// ═══════════════════════════════════════════════════════════════════════════════
// [4] DELL BLOATWARE SERVICES
// ═══════════════════════════════════════════════════════════════════════════════

static void StopDellBloatwareServices(List<string> restore)
{
    SectionHeader("STOP & DISABLE DELL BLOATWARE SERVICES");

    string[] services =
    [
        "DellClientManagementService",
        "DellTechHub",
        "Dell SupportAssist Remediation",
        "SupportAssistAgent",
        "DellPairService",
        "DellTrustedDevice",
    ];

    StopAndDisableServices(services, restore);
}

// ═══════════════════════════════════════════════════════════════════════════════
// [5] REMOTE ACCESS SERVICES
// ═══════════════════════════════════════════════════════════════════════════════

static void StopRemoteAccessServices(List<string> restore)
{
    SectionHeader("STOP & DISABLE REDUNDANT REMOTE ACCESS SERVICES");
    WriteColor("  Note: ScreenConnect is kept (likely primary remote tool).", ConsoleColor.DarkGray);

    string[] services =
    [
        "SplashtopRemoteService",
        "TeamViewer",
        "SAAZappr",
        "SAAZDPMACTL",
        "SAAZRemoteSupport",
        "SAAZScheduler",
        "SAAZServerPlus",
        "SAAZWatchDog",
        "ITSPlatform",
        "ITSPlatformManager",
    ];

    StopAndDisableServices(services, restore);
}

static void StopAndDisableServices(string[] serviceNames, List<string> restore)
{
    foreach (var name in serviceNames)
    {
        ServiceController? svc = null;
        try
        {
            svc = new ServiceController(name);
            _ = svc.Status; // triggers exception if service doesn't exist
        }
        catch
        {
            WriteColor($"  [{name}] — not found, skipping.", ConsoleColor.DarkGray);
            svc?.Dispose();
            continue;
        }

        try
        {
            string status = svc.Status.ToString();
            var startType = GetServiceStartType(name);
            Console.WriteLine($"  [{name}] Status={status}, StartType={startType}");

            if (svc.Status == ServiceControllerStatus.Stopped &&
                startType.Equals("Disabled", StringComparison.OrdinalIgnoreCase))
            {
                WriteColor($"    Already stopped and disabled.", ConsoleColor.DarkGray);
                continue;
            }

            if (!Confirm($"Stop & disable '{name}'?"))
                continue;

            // Stop if running
            if (svc.Status != ServiceControllerStatus.Stopped)
            {
                try
                {
                    svc.Stop();
                    svc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
                    WriteColor($"    Stopped.", ConsoleColor.Green);
                }
                catch (Exception ex)
                {
                    WriteColor($"    Could not stop: {ex.Message}", ConsoleColor.Red);
                }
            }

            // Disable via sc.exe (most reliable cross-version method)
            var psi = new ProcessStartInfo("sc.exe", $"config \"{name}\" start= disabled")
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var proc = Process.Start(psi)!;
            proc.StandardOutput.ReadToEnd(); // drain stdout to avoid deadlock
            proc.WaitForExit(10000);
            if (proc.ExitCode == 0)
            {
                WriteColor($"    Disabled.", ConsoleColor.Green);
                restore.Add($"# Restore service: {name} (was {startType})");
                string scStartType = startType switch
                {
                    "Automatic" => "auto",
                    "Manual" => "demand",
                    _ => "demand"
                };
                restore.Add($"sc.exe config \"{name}\" start= {scStartType}");
                restore.Add($"# sc.exe start \"{name}\"  # Uncomment to also restart");
                restore.Add("");
            }
            else
            {
                WriteColor($"    Failed to disable (sc.exe exit code {proc.ExitCode}).", ConsoleColor.Red);
            }
        }
        catch (Exception ex)
        {
            WriteColor($"  [{name}] Error: {ex.Message}", ConsoleColor.Red);
        }
        finally
        {
            svc.Dispose();
        }
    }
}

static string GetServiceStartType(string serviceName)
{
    try
    {
        using var key = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{serviceName}");
        if (key != null)
        {
            int start = (int)(key.GetValue("Start") ?? -1);
            return start switch
            {
                0 => "Boot",
                1 => "System",
                2 => "Automatic",
                3 => "Manual",
                4 => "Disabled",
                _ => $"Unknown({start})"
            };
        }
    }
    catch { }
    return "Unknown";
}

// ═══════════════════════════════════════════════════════════════════════════════
// [6] SCHEDULED TASKS
// ═══════════════════════════════════════════════════════════════════════════════

static void DisableScheduledTasks(List<string> restore)
{
    SectionHeader("DISABLE UNNECESSARY SCHEDULED TASKS");

    string[] tasks =
    [
        @"\MSIAfterburner",
        @"\Microsoft\Office\Office Background Push Maintenance",
    ];

    // Also find any VS UpdateConfiguration tasks
    string[] wildcardPrefixes =
    [
        @"\Microsoft\VisualStudio\Updates\UpdateConfiguration_",
    ];

    var allTasks = new List<string>(tasks);

    // Expand wildcards by enumerating
    try
    {
        var psi = new ProcessStartInfo("schtasks.exe", "/query /fo CSV /nh")
        {
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var proc = Process.Start(psi)!;
        string output = proc.StandardOutput.ReadToEnd();
        proc.WaitForExit();

        foreach (string line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            // CSV format: "TaskName","Next Run Time","Status"
            string taskName = line.Split(',')[0].Trim('"', ' ', '\r');
            foreach (var prefix in wildcardPrefixes)
            {
                if (taskName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    && !allTasks.Contains(taskName))
                {
                    allTasks.Add(taskName);
                }
            }
        }
    }
    catch (Exception ex)
    {
        WriteColor($"  Warning: could not enumerate tasks: {ex.Message}", ConsoleColor.Yellow);
    }

    foreach (var taskPath in allTasks)
    {
        Console.WriteLine($"  Task: {taskPath}");

        if (!Confirm($"Disable task '{taskPath}'?"))
            continue;

        try
        {
            var psi = new ProcessStartInfo("schtasks.exe", $"/change /tn \"{taskPath}\" /disable")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var proc = Process.Start(psi)!;
            // Read stderr on a background thread to avoid deadlock when a pipe buffer fills
            string stderr = "";
            var stderrThread = new Thread(() => stderr = proc.StandardError.ReadToEnd());
            stderrThread.Start();
            string stdout = proc.StandardOutput.ReadToEnd();
            stderrThread.Join();
            proc.WaitForExit();

            if (proc.ExitCode == 0)
            {
                WriteColor($"    Disabled.", ConsoleColor.Green);
                restore.Add($"# Re-enable scheduled task: {taskPath}");
                restore.Add($"schtasks.exe /change /tn \"{taskPath}\" /enable");
                restore.Add("");
            }
            else
            {
                string msg = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
                WriteColor($"    Failed: {msg.Trim()}", ConsoleColor.Red);
            }
        }
        catch (Exception ex)
        {
            WriteColor($"    Error: {ex.Message}", ConsoleColor.Red);
        }
    }
}

// ═══════════════════════════════════════════════════════════════════════════════
// [7] KILL RESOURCE-HEAVY PROCESSES
// ═══════════════════════════════════════════════════════════════════════════════

static void KillResourceHeavyProcesses(List<string> restore)
{
    SectionHeader("KILL RESOURCE-HEAVY NON-ESSENTIAL PROCESSES");

    string[] targets =
    [
        "PhoneExperienceHost",
        "LM Studio",
        "SDXHelper",
    ];

    // Exact-name targets
    foreach (var name in targets)
        KillProcessesByName(name);

    // Wildcard: WhatsApp*
    foreach (var p in Process.GetProcesses())
    {
        try
        {
            if (p.ProcessName.StartsWith("WhatsApp", StringComparison.OrdinalIgnoreCase))
                KillSingleProcess(p);
            else
                p.Dispose();
        }
        catch { p.Dispose(); }
    }
}

static void KillProcessesByName(string name)
{
    var procs = Process.GetProcessesByName(name);
    if (procs.Length == 0)
    {
        // Also try without spaces (process names sometimes differ)
        procs = Process.GetProcessesByName(name.Replace(" ", ""));
    }
    if (procs.Length == 0)
    {
        WriteColor($"  [{name}] — not running.", ConsoleColor.DarkGray);
        return;
    }

    foreach (var p in procs)
        KillSingleProcess(p);
}

static void KillSingleProcess(Process p)
{
    try
    {
        long mem = p.WorkingSet64;
        string name = p.ProcessName;
        if (Confirm($"Kill {name} PID {p.Id} ({FormatBytes(mem)})?"))
        {
            p.Kill(entireProcessTree: true);
            p.WaitForExit(5000);
            WriteColor($"    Killed {name} PID {p.Id} — freed {FormatBytes(mem)}", ConsoleColor.Green);
        }
    }
    catch (Exception ex)
    {
        WriteColor($"    Error: {ex.Message}", ConsoleColor.Red);
    }
    finally
    {
        p.Dispose();
    }
}

// ═══════════════════════════════════════════════════════════════════════════════
// [8] STARTUP ENTRIES
// ═══════════════════════════════════════════════════════════════════════════════

static void DisableStartupEntries(List<string> restore)
{
    SectionHeader("DISABLE UNNECESSARY STARTUP ENTRIES");

    // HKCU entries
    DisableStartupFromRegistry(
        Registry.CurrentUser,
        @"Software\Microsoft\Windows\CurrentVersion\Run",
        "HKCU",
        [
            ("LM Studio", "electron.app.LM Studio"),
            ("Logi Tune", "Logi Tune"),
        ],
        ["GoogleChromeAutoLaunch_"],
        restore
    );

    // HKLM entries (system-wide)
    DisableStartupFromRegistry(
        Registry.LocalMachine,
        @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
        "HKLM",
        [],
        [],
        restore
    );
}

static void DisableStartupFromRegistry(
    RegistryKey hive,
    string subKeyPath,
    string hiveName,
    (string displayName, string valueName)[] exactEntries,
    string[] prefixes,
    List<string> restore)
{
    try
    {
        using var key = hive.OpenSubKey(subKeyPath, writable: true);
        if (key == null)
        {
            WriteColor($"  Registry key {hiveName}\\{subKeyPath} not found.", ConsoleColor.Yellow);
            return;
        }

        var valueNames = key.GetValueNames();
        var toProcess = new List<(string valueName, string displayName)>();

        // Exact matches
        foreach (var (display, target) in exactEntries)
        {
            var match = valueNames.FirstOrDefault(v =>
                v.Equals(target, StringComparison.OrdinalIgnoreCase) ||
                v.Contains(target, StringComparison.OrdinalIgnoreCase));
            if (match != null)
                toProcess.Add((match, display));
            else
                WriteColor($"  [{display}] — not found in {hiveName} startup, skipping.", ConsoleColor.DarkGray);
        }

        // Prefix matches (e.g., GoogleChromeAutoLaunch_*)
        foreach (var prefix in prefixes)
        {
            foreach (var vn in valueNames)
            {
                if (vn.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    toProcess.Add((vn, vn));
            }
        }

        foreach (var (valueName, displayName) in toProcess)
        {
            var currentValue = key.GetValue(valueName);
            Console.WriteLine($"  [{displayName}] Value: {currentValue}");

            if (!Confirm($"Remove startup entry '{displayName}'?"))
                continue;

            try
            {
                key.DeleteValue(valueName);
                WriteColor($"    Removed.", ConsoleColor.Green);
                string escapedValue = currentValue?.ToString()?.Replace("'", "''") ?? "";
                restore.Add($"# Restore startup entry: {displayName}");
                restore.Add($"New-ItemProperty -Path '{hiveName}:\\{subKeyPath}' -Name '{valueName}' -Value '{escapedValue}' -PropertyType String -Force");
                restore.Add("");
            }
            catch (Exception ex)
            {
                WriteColor($"    Failed: {ex.Message}", ConsoleColor.Red);
            }
        }
    }
    catch (Exception ex)
    {
        WriteColor($"  Error accessing {hiveName}\\{subKeyPath}: {ex.Message}", ConsoleColor.Red);
    }
}

// ═══════════════════════════════════════════════════════════════════════════════
// [9] DEFENDER EXCLUSIONS
// ═══════════════════════════════════════════════════════════════════════════════

static async Task AddDefenderExclusions(List<string> restore)
{
    SectionHeader("ADD WINDOWS DEFENDER EXCLUSIONS FOR DEV PATHS");

    string[] paths =
    [
        @"C:\Program Files\dotnet\",
        @"C:\Program Files\Microsoft Visual Studio\",
        @"C:\Users\DVeksler\.nuget\",
        @"C:\Users\DVeksler\source\",
    ];

    // Gather existing exclusions so we don't duplicate
    HashSet<string> existing = [];
    try
    {
        var (output, _) = await RunProcess("powershell", "-NoProfile -Command \"(Get-MpPreference).ExclusionPath -join '|'\"");
        foreach (var p in output.Trim().Split('|', StringSplitOptions.RemoveEmptyEntries))
            existing.Add(p.Trim().TrimEnd('\\'));
    }
    catch { }

    foreach (var path in paths)
    {
        string normalized = path.TrimEnd('\\');
        if (existing.Contains(normalized))
        {
            WriteColor($"  [{path}] — already excluded.", ConsoleColor.DarkGray);
            continue;
        }

        if (!Directory.Exists(path.TrimEnd('\\')))
        {
            WriteColor($"  [{path}] — path does not exist, skipping.", ConsoleColor.Yellow);
            continue;
        }

        Console.WriteLine($"  Path: {path}");
        if (!Confirm($"Add Defender exclusion for '{path}'?"))
            continue;

        var (_, exitCode) = await RunProcess("powershell",
            $"-NoProfile -Command \"Add-MpPreference -ExclusionPath '{path}'\"");

        if (exitCode == 0)
        {
            WriteColor($"    Exclusion added.", ConsoleColor.Green);
            restore.Add($"# Remove Defender exclusion: {path}");
            restore.Add($"Remove-MpPreference -ExclusionPath '{path}'");
            restore.Add("");
        }
        else
        {
            WriteColor($"    Failed to add exclusion.", ConsoleColor.Red);
        }
    }
}
