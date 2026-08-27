// vybe-test: csharp/csharp_linq_deferred_execution/linq_select_does_not_run_until_foreach_starts
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using static __Harness;
using System.Linq;

int sideEffects = 0;
var query = new[] { 1, 2 }
.Select(x => { sideEffects++; return x; });
__P((sideEffects).ToString());
foreach (var _ in query) { }
__P((sideEffects).ToString());
__Check("0\n2");

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
