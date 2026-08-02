// vybe-test: csharp/csharp_collection_initializer_syntax/readonly_struct_initializer_sets_init_only_properties
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Token {
    public string Value { get; init; }
}
var token = new Token { Value = "abc" };
__Check((token.Value).ToString(), "abc");
