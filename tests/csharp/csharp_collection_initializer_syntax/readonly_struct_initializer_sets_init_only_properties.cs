// vybe-test: csharp/csharp_collection_initializer_syntax/readonly_struct_initializer_sets_init_only_properties
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Token {
    public string Value { get; init; }
}
var token = new Token { Value = "abc" };
__P((token.Value).ToString());
__Check("abc");
