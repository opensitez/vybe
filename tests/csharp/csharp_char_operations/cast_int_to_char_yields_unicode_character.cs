// vybe-test: csharp/csharp_char_operations/cast_int_to_char_yields_unicode_character
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(((char)65).ToString(), "A");
