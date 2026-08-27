// vybe-test: csharp/csharp_linq_materialization/linq_count_executes_predicate_immediately
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

using static __Harness;
using System.Linq;

int checks = 0;
int total = new[] { 1, 2, 3 }
.Count(x => { checks++; return x > 1; });
__P((total).ToString());
__P((checks).ToString());
__Check("2\n3");

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
