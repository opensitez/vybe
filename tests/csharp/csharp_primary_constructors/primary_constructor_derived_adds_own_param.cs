// vybe-test: csharp/csharp_primary_constructors/primary_constructor_derived_adds_own_param
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

__P((new Extra(2, 5).Y).ToString());
__Check("5");

class Base(int x) { public int X => x; }

class Extra(int x, int y) : Base(x) { public int Y => y; }

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
