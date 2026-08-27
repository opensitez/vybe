// vybe-test: csharp/csharp_primary_constructors/primary_constructor_struct_copies_params_to_fields
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

var p = new Point(3, 4);
__P((p.X).ToString());
__P((p.Y).ToString());
__Check("3\n4");

struct Point(int x, int y) {
    public int X = x;
    public int Y = y;
}

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
