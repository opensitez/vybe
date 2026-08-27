// vybe-test: csharp/csharp_primary_constructors/primary_constructor_record_with_extra_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

__P((new Point(2, 3).Sum()).ToString());
__Check("5");

record Point(int X, int Y) { public int Sum() => X + Y; }

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
