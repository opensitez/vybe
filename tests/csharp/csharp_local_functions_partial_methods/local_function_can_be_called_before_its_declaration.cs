// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_be_called_before_its_declaration
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

using static __Harness;

int LocalHelper(int x) => x * 2;
__P(LocalHelper(10).ToString());
__Check("20");
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
