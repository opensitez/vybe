// vybe-test: csharp/csharp_linq_materialization/linq_min_returns_smallest_element_by_default_comparer
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

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
__P((new[] { 3, 9, 4 }.Min()).ToString());
__Check("3");
