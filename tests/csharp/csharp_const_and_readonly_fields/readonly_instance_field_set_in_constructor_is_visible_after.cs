// vybe-test: csharp/csharp_const_and_readonly_fields/readonly_instance_field_set_in_constructor_is_visible_after
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

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
__Check((new Token("key").Value).ToString(), "key");
