// vybe-test: csharp/csharp_decimal_semantics/decimal_increment_mutates_storage_in_place
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

using static __Harness;

decimal tally = 2.5m;
tally++;
__P((tally).ToString());
__Check("3.5");

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
