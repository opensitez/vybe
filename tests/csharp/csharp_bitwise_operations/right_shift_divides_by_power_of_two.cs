// vybe-test: csharp/csharp_bitwise_operations/right_shift_divides_by_power_of_two
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

using static __Harness;

__P((64 >> 3).ToString());
__Check("8");

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
