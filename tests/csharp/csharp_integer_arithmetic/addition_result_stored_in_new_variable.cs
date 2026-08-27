// vybe-test: csharp/csharp_integer_arithmetic/addition_result_stored_in_new_variable
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

using static __Harness;

int sum = 14 + 6;
__P((sum).ToString());
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
