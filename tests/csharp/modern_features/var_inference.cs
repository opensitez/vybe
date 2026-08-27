// vybe-test: csharp/modern_features/var_inference
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var x = 42;
var s = "hello";
var list = new List<int> { 1, 2, 3 }
;
__P((x.GetType().Name).ToString());
__P((s.GetType().Name).ToString());
__P((list.Count).ToString());
__Check("Int32\nString\n3");

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
