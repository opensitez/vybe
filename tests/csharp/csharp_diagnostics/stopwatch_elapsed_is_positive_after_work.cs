// vybe-test: csharp/csharp_diagnostics/stopwatch_elapsed_is_positive_after_work
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
int s=0; for(int i=0;i<10000;i++) s+=i;
sw.Stop();
__P((sw.ElapsedMilliseconds>=0).ToString());
__Check("True");
