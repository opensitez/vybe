// vybe-test: csharp/csharp_linq_deferred_execution/linq_orderby_defers_sort_until_materialization
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
int comparisons = 0;
var query = new[] { 3, 1, 2 }.OrderBy(x => { comparisons++; return x; });
__P((comparisons).ToString());
foreach (var value in query) __P((value).ToString());
__Check("0\n1\n2\n3");
