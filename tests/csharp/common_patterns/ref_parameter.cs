// vybe-test: csharp/common_patterns/ref_parameter
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

int val = 5;
Ops.Double(ref val);
__P((val).ToString());
Ops.Double(ref val);
__P((val).ToString());
__Check("10\n20");

class Ops {
    public static void Double(ref int x) { x *= 2; }
}

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
