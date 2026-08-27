// vybe-test: csharp/basics/typed_declaration
// origin: languages/csharp/tests/csharp/test_basics.rs

using static __Harness;

int x = 5;
double y = 3.14;
__P((x).ToString());
__P((y).ToString());
__Check("5\n3.14");

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
