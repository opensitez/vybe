// vybe-test: csharp/csharp_unsafe_pointers/fixed_statement_pins_array_for_pointer_arithmetic
// origin: languages/csharp/tests/csharp/test_csharp_unsafe_pointers.rs

using static __Harness;

int val = 100;
__P(val.ToString());
__Check("100");
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
