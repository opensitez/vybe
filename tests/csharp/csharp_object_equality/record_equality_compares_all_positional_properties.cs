// vybe-test: csharp/csharp_object_equality/record_equality_compares_all_positional_properties
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

using static __Harness;

var a = new Point(1, 2);
var b = new Point(1, 2);
var c = new Point(1, 3);
__P((a.Equals(b)).ToString());
__P((a.Equals(c)).ToString());
__Check("True\nFalse");

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
