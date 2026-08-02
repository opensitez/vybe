// vybe-test: csharp/csharp_structs_value_semantics/struct_can_override_to_string_for_custom_output
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Token { public int Value; public override string ToString() { return "T:" + Value; } } __Check((new Token { Value = 7 }).ToString(), "T:7");
