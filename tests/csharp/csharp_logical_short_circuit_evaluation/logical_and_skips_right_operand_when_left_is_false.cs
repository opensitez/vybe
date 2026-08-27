// vybe-test: csharp/csharp_logical_short_circuit_evaluation/logical_and_skips_right_operand_when_left_is_false
// origin: languages/csharp/tests/csharp/test_csharp_logical_short_circuit_evaluation.rs

using static __Harness;

int calls = 0;
bool Right() { calls++; return true; }
bool result = false && Right();
__P((result ? "T" : "F").ToString());
__P((calls).ToString());
__Check("F\n0");

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
