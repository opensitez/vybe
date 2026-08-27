// vybe-test: csharp/csharp_bitwise_operations/compound_bitwise_and_assign_updates_in_place
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

using static __Harness;

int x = 0b1111;
x &= 0b0101;
__P((x).ToString());
__Check("5");

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
