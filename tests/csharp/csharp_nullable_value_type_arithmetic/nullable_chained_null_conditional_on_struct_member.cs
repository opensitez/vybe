// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_chained_null_conditional_on_struct_member
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

using static __Harness;

Point? location = new Point { X = 2, Y = 3 }
;
__P((location?.X).ToString());
location = null;
__P((location?.X ?? -1).ToString());
__Check("2\n-1");

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
