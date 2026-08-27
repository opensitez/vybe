// vybe-test: csharp/csharp_expression_bodied_members/expr_method_property_indexer_combined
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var c = new Cache();
c[0] = 1;
c[1] = 2;
c[2] = 3;
__P((c.Sum()).ToString());
__Check("6");

class Cache { int[] buf = { 0, 0, 0 }; public int this[int i] { get => buf[i]; set => buf[i] = value; } public int Sum() => buf[0] + buf[1] + buf[2]; }

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
