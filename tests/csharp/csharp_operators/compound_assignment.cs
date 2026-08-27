// vybe-test: csharp/csharp_operators/compound_assignment
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

using static __Harness;

int x = 10;
x += 5;
__P((x).ToString());
x -= 3;
__P((x).ToString());
x *= 2;
__P((x).ToString());
x /= 4;
__P((x).ToString());
x %= 5;
__P((x).ToString());
__Check("15\n12\n24\n6\n1");

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
