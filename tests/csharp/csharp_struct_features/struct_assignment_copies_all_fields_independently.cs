// vybe-test: csharp/csharp_struct_features/struct_assignment_copies_all_fields_independently
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

using static __Harness;

var a = new Point { X=1, Y=2 }
;
var b = a;
b.X = 99;
__P((a.X).ToString());
__Check("1");

struct Point { public int X, Y; }

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
