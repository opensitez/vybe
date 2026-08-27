// vybe-test: csharp/csharp_collection_initializer_syntax/readonly_struct_initializer_sets_init_only_properties
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;

var token = new Token { Value = "abc" }
;
__P((token.Value).ToString());
__Check("abc");

readonly struct Token {
    public string Value { get; init; }
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
