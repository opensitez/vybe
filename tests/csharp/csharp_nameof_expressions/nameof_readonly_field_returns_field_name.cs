// vybe-test: csharp/csharp_nameof_expressions/nameof_readonly_field_returns_field_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Token{public readonly string Value="x";} __Check((nameof(Token.Value)).ToString(), "Value");
