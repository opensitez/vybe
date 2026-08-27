// vybe-test: csharp/csharp_array_apis/array_for_each_invokes_action_for_each_item
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

using static __Harness;

var values = new[] { 3, 4 }
;
System.Array.ForEach(values, value => __P((value * 2).ToString()));
__Check("6\n8");

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
