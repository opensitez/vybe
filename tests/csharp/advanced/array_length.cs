// vybe-test: csharp/advanced/array_length
// origin: languages/csharp/tests/csharp/test_advanced.rs

using static __Harness;

var arr = new int[] { 1, 2, 3 }
;
__P((arr[0]).ToString());
__P((arr[2]).ToString());
__Check("1\n3");

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
