// vybe-test: csharp/csharp_structs_value_semantics/struct_can_have_readonly_field_initialized_by_constructor
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

__P((new Token(5).Value).ToString());
__Check("5");

struct Token { public readonly int Value; public Token(int value) { Value = value; } }

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
