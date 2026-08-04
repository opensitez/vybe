// vybe-test: csharp/csharp_linq_deferred_execution/linq_where_filter_runs_only_during_enumeration
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

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

using System.Linq;
int checks = 0;
var query = new[] { 1, 2, 3, 4 }.Where(x => { checks++; return x % 2 == 0; });
__P((checks).ToString());
foreach (var value in query) __P((value).ToString());
__P((checks).ToString());
__Check("0\n2\n4\n4");
