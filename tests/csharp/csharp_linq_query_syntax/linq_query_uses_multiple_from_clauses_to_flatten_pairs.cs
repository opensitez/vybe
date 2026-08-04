// vybe-test: csharp/csharp_linq_query_syntax/linq_query_uses_multiple_from_clauses_to_flatten_pairs
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

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
var prefixes = new[] { "A", "B" };
var suffixes = new[] { 1, 2, 3 };
var query = from prefix in prefixes
            from suffix in suffixes
            where suffix != 2
            select prefix + suffix;
foreach (var item in query) __P((item).ToString());
__Check("A1\nA3\nB1\nB3");
