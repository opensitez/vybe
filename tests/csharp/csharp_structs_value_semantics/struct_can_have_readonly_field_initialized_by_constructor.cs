// vybe-test: csharp/csharp_structs_value_semantics/struct_can_have_readonly_field_initialized_by_constructor
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Token { public readonly int Value; public Token(int value) { Value = value; } } __Check((new Token(5).Value).ToString(), "5");
