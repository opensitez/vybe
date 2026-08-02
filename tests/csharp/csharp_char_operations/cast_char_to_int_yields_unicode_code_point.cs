// vybe-test: csharp/csharp_char_operations/cast_char_to_int_yields_unicode_code_point
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(((int)'A').ToString(), "65");
