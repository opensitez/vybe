// vybe-test: csharp/csharp_delegate_types/action_delegate_calls_void_method_with_no_args
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

using static __Harness;

System.Action greet = () => __P(("hi").ToString());
greet();
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
