// vybe-test: csharp/csharp_diagnostics/stopwatch_reset_clears_elapsed
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
System.Threading.Thread.Sleep(5);
sw.Stop();
sw.Reset();
__P((sw.Elapsed==System.TimeSpan.Zero).ToString());
__Check("True");
