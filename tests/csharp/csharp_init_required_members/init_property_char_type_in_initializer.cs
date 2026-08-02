// vybe-test: csharp/csharp_init_required_members/init_property_char_type_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Token { public char Symbol { get; init; } = 'a'; }
var t = new Token { Symbol = 'z' };
__Check((t.Symbol).ToString(), "z");
