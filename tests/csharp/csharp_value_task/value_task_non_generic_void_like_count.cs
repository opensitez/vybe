// vybe-test: csharp/csharp_value_task/value_task_non_generic_void_like_count
// origin: languages/csharp/tests/csharp/test_csharp_value_task.rs

using static __Harness;

__P("Valid_value_task_non_generic_void_like_count");
__Check("Valid_value_task_non_generic_void_like_count");
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
