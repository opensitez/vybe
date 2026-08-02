// vybe-test: csharp/csharp_initialization_order/readonly_instance_field_can_be_set_in_constructor_but_not_after
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Token {
    public readonly string Value;
    public Token(string value) { Value = value; }
}
var token = new Token("abc");
__Check((token.Value).ToString(), "abc");
