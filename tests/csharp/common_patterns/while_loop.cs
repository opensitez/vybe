// vybe-test: csharp/common_patterns/while_loop
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

int x = 1;
while (x <= 16) {
    __P((x).ToString());
    x *= 2;
}
__Check("1\n2\n4\n8\n16");

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
