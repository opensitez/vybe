// vybe-test: csharp/csharp_record_types/records_generated_equals_compares_all_properties
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

using static __Harness;

var a = new Point(1,2);
var b = new Point(1,2);
var c = new Point(1,3);
__P((a == b).ToString());
__P((a == c).ToString());
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
