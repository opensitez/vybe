// vybe-test: csharp/csharp_const_and_readonly_fields/readonly_instance_field_set_in_constructor_is_visible_after
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

using static __Harness;

__P((new Token("key").Value).ToString());
__Check("key");

class Token {
    public readonly string Value;
    public Token(string value) { Value = value; }
}

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
