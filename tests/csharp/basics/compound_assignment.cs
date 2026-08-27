// vybe-test: csharp/basics/compound_assignment
// origin: languages/csharp/tests/csharp/test_basics.rs

using static __Harness;

var x = 10;
x += 5;
x -= 3;
x *= 2;
__P((x).ToString());
__Check("24");

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
