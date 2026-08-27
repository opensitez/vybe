// vybe-test: csharp/csharp_primary_constructors/primary_constructor_struct_copy_preserves_field_values
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

var p = new Pair(2, 3);
var q = p;
__P((q.A + q.B).ToString());
__Check("5");

struct Pair(int a, int b) { public int A = a; public int B = b; }

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
