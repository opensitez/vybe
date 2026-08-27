// vybe-test: csharp/csharp_delegate_types/func_t_result_returns_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

using static __Harness;

System.Func<int,int> square = x => x * x;
__P((square(4)).ToString());
__Check("16");

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
