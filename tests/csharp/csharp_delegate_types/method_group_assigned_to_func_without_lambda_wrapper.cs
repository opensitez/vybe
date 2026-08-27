// vybe-test: csharp/csharp_delegate_types/method_group_assigned_to_func_without_lambda_wrapper
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

using static __Harness;

System.Func<string,int> len = s => s.Length;
__P((len("test")).ToString());
__Check("4");

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
