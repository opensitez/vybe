// vybe-test: csharp/csharp_logical_short_circuit_evaluation/and_short_circuits_before_or_evaluates_fallback_operand
// origin: languages/csharp/tests/csharp/test_csharp_logical_short_circuit_evaluation.rs

using static __Harness;

int trace = 0;
bool A() { trace++; return false; }
bool B() { trace++; return true; }
bool C() { trace++; return true; }
bool value = A() && B() || C();
__P((value ? "T" : "F").ToString());
__P((trace).ToString());
__Check("T\n2");

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
