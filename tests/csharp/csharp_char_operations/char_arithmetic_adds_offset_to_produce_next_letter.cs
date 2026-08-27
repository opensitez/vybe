// vybe-test: csharp/csharp_char_operations/char_arithmetic_adds_offset_to_produce_next_letter
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

using static __Harness;

char c = (char)('A' + 2);
__P((c).ToString());
__Check("C");

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
