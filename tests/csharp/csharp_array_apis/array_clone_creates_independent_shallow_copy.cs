// vybe-test: csharp/csharp_array_apis/array_clone_creates_independent_shallow_copy
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

using static __Harness;

var source = new[] { 1, 2 }
;
var clone = (int[])source.Clone();
clone[0] = 9;
__P((source[0]).ToString());
__P((clone[0]).ToString());
__Check("1\n9");

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
