// vybe-test: csharp/csharp_init_required_members/required_property_constructor_then_read
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Token { public required string Value { get; set; } public Token() { Value = "init"; } }
var t = new Token();
__Check((t.Value).ToString(), "init");
