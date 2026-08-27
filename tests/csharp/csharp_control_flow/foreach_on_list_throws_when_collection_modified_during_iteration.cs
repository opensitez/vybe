// vybe-test: csharp/csharp_control_flow/foreach_on_list_throws_when_collection_modified_during_iteration
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

using static __Harness;

var items = new System.Collections.Generic.List<int> { 1, 2 }
;
string outcome = "ok";
try {
    foreach (var item in items) {
        items.RemoveAt(0);
    }
}
catch (System.InvalidOperationException) {
    outcome = "invalid";
}
__P((outcome).ToString());
__Check("invalid");

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
