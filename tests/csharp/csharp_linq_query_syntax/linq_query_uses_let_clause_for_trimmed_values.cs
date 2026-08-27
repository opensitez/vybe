// vybe-test: csharp/csharp_linq_query_syntax/linq_query_uses_let_clause_for_trimmed_values
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using static __Harness;
using System.Linq;

var raw = new[] { "  alpha  ", " beta", "gamma " }
;
var query = from value in raw
            let trimmed = value.Trim()
            select trimmed + ":" + trimmed.Length;
foreach (var item in query) __P((item).ToString());
__Check("alpha:5\nbeta:4\ngamma:5");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
