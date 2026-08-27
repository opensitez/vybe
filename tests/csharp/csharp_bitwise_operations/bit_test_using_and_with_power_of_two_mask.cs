// vybe-test: csharp/csharp_bitwise_operations/bit_test_using_and_with_power_of_two_mask
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

using static __Harness;

int flags = 0b1010;
__P(((flags & 0b0010) != 0).ToString());
__Check("True");

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
