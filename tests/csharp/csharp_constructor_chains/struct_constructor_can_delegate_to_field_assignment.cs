// vybe-test: csharp/csharp_constructor_chains/struct_constructor_can_delegate_to_field_assignment
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

var pair = new Pair(2, 8);
__P((pair.Left + pair.Right).ToString());
__Check("10");

struct Pair { public int Left; public int Right; public Pair(int left, int right) { Left = left; Right = right; } }

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
