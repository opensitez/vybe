// vybe-test: csharp/csharp_linq_deferred_execution/linq_second_enumeration_reexecutes_select_projection
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using static __Harness;
using System.Linq;

int projections = 0;
var query = new[] { 5 }
.Select(x => { projections++; return x + 1; });
__P((query.First()).ToString());
__P((projections).ToString());
__P((query.First()).ToString());
__P((projections).ToString());
__Check("6\n1\n6\n2");

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
