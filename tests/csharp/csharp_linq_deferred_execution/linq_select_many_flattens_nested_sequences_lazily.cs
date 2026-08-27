// vybe-test: csharp/csharp_linq_deferred_execution/linq_select_many_flattens_nested_sequences_lazily
// origin: languages/csharp/tests/csharp/test_csharp_linq_deferred_execution.rs

using static __Harness;
using System.Linq;

var query = new[] { "ab", "c" }
.SelectMany(word => word.Select(ch => ch));
foreach (var ch in query) __P((ch).ToString());
__Check("a\nb\nc");

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
