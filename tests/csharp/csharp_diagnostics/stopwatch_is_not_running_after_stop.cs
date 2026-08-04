// vybe-test: csharp/csharp_diagnostics/stopwatch_is_not_running_after_stop
// origin: languages/csharp/tests/csharp/test_csharp_diagnostics.rs

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

var sw=System.Diagnostics.Stopwatch.StartNew();
sw.Stop();
__P((sw.IsRunning).ToString());
__Check("False");
