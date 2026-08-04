// vybe-test: csharp/linq_lambdas/linq_orderby_thenby
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

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

var names = new List<string> { "Charlie", "Alice", "Bob", "Alice" };
var sorted = names.OrderBy(n => n).ToList();
foreach (var n in sorted) __P((n).ToString());
__Check("Alice\nAlice\nBob\nCharlie");
