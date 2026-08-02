// vybe-test: csharp/csharp_char_type_semantics/char_comparison_uses_numeric_code_unit_ordering
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(('A' < 'B').ToString(), "True");
