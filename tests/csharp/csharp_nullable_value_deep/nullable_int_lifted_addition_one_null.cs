// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_lifted_addition_one_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

using static __Harness;

int? a=10;
int? b=null;
__P(((a+b).HasValue).ToString());
__Check("False");

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
