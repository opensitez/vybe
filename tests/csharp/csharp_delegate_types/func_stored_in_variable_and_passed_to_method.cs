// vybe-test: csharp/csharp_delegate_types/func_stored_in_variable_and_passed_to_method
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

using static __Harness;

int Apply(System.Func<int,int> f, int v) => f(v);
__P((Apply(x => x + 1, 9)).ToString());
__Check("10");

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
