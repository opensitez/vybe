// vybe-test: csharp/csharp_null_propagation/null_coalescing_assignment_sets_value_when_variable_is_null
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

using static __Harness;

string value = null;
value ??= "set";
__P((value).ToString());
__Check("set");

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
