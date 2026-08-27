// vybe-test: csharp/csharp_integer_arithmetic/compound_assignment_chain_updates_final_value
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

using static __Harness;

int value = 10;
value += 5;
value *= 2;
value -= 8;
__P((value).ToString());
__Check("22");

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
