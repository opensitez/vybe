// vybe-test: csharp/csharp_iterators/yield_break_stops_iteration_early
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

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

System.Collections.Generic.IEnumerable<int> Gen() {
    yield return 1;
    yield break;
    yield return 2;
}
int count = 0;
foreach(var _ in Gen()) count++;
__P((count).ToString());
__Check("1");
