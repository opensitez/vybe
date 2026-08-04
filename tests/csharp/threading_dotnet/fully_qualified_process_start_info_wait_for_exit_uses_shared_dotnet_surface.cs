// vybe-test: csharp/threading_dotnet/fully_qualified_process_start_info_wait_for_exit_uses_shared_dotnet_surface
// origin: languages/csharp/tests/csharp/test_threading_dotnet.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var si = new System.Diagnostics.ProcessStartInfo("/usr/bin/test", "hello = hello");
        var p = System.Diagnostics.Process.Start(si);
        p.WaitForExit();
        __P((p.ExitCode).ToString());
__Check("0");
