// vybe-test: csharp/modern_features/record_tostring
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var p = new Point(3, 4);
__P((p).ToString());
__Check("Point { X = 3, Y = 4 }");

record Point(int X, int Y);

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
