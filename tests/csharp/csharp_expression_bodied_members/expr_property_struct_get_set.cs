// vybe-test: csharp/csharp_expression_bodied_members/expr_property_struct_get_set
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var s = new Slot();
s.N = 7;
__P((s.N).ToString());
__Check("7");

struct Slot { int _n; public int N { get => _n; set => _n = value; } }

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
