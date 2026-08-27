// vybe-test: csharp/csharp_control_flow/while_loop
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

using static __Harness;

int i = 0;
while (i < 3) {
    __P((i).ToString());
    i++;
}
__Check("0\n1\n2");

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
