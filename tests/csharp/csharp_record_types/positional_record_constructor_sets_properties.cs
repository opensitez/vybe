// vybe-test: csharp/csharp_record_types/positional_record_constructor_sets_properties
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

using static __Harness;

var p = new Point(3,4);
__P((p.X).ToString());
__P((p.Y).ToString());
__Check("3\n4");

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
