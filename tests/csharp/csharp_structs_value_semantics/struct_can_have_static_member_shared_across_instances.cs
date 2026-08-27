// vybe-test: csharp/csharp_structs_value_semantics/struct_can_have_static_member_shared_across_instances
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

new Token(1);
new Token(2);
__P((Token.Count).ToString());
__Check("2");

struct Token { public static int Count; public Token(int _) { Count++; } }

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
