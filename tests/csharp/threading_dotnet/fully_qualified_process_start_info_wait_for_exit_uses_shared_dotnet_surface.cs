// vybe-test: csharp/threading_dotnet/fully_qualified_process_start_info_wait_for_exit_uses_shared_dotnet_surface
// origin: languages/csharp/tests/csharp/test_threading_dotnet.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var si = new System.Diagnostics.ProcessStartInfo("/usr/bin/test", "hello = hello");
        var p = System.Diagnostics.Process.Start(si);
        p.WaitForExit();
        __Check((p.ExitCode).ToString(), "0");
