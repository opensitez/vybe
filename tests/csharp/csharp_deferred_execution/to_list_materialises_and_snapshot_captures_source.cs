// vybe-test: csharp/csharp_deferred_execution/to_list_materialises_and_snapshot_captures_source
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

using static __Harness;

var source=new System.Collections.Generic.List<int>{1,2,3}
;
var snapshot=source.ToList();
source.Add(4);
__P((snapshot.Count).ToString());
__Check("3");

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
