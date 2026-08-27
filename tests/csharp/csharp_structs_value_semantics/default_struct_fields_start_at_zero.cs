// vybe-test: csharp/csharp_structs_value_semantics/default_struct_fields_start_at_zero
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

var point = new Point();
__P((point.X).ToString());
__P((point.Y).ToString());
__Check("0\n0");

struct Point { public int X; public int Y; }

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
