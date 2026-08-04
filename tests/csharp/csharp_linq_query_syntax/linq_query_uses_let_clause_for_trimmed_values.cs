// vybe-test: csharp/csharp_linq_query_syntax/linq_query_uses_let_clause_for_trimmed_values
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
var raw = new[] { "  alpha  ", " beta", "gamma " };
var query = from value in raw
            let trimmed = value.Trim()
            select trimmed + ":" + trimmed.Length;
foreach (var item in query) __P((item).ToString());
__Check("alpha:5\nbeta:4\ngamma:5");
