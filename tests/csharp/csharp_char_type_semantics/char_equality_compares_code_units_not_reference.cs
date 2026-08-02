// vybe-test: csharp/csharp_char_type_semantics/char_equality_compares_code_units_not_reference
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char left = 'Z';
char right = 'Z';
__Check((left == right).ToString(), "True");
