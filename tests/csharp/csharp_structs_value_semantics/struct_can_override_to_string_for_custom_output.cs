// vybe-test: csharp/csharp_structs_value_semantics/struct_can_override_to_string_for_custom_output
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

__P((new Token { Value = 7 }).ToString());
__Check("T:7");

struct Token { public int Value; public override string ToString() { return "T:" + Value; } }

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
