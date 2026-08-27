// vybe-test: csharp/csharp_delegates/action_no_params
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

using static __Harness;

Action sayHi = () => __P(("hi").ToString());
sayHi();
__Check("hi");

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
