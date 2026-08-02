// vybe-test: csharp/csharp_char_type_semantics/char_array_holds_sequence_of_code_units
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char[] letters = { 'a', 'b', 'c' };
__Check((letters[2]).ToString(), "c");
