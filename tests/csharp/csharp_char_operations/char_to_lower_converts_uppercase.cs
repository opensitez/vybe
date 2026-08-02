// vybe-test: csharp/csharp_char_operations/char_to_lower_converts_uppercase
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((char.ToLower('Z')).ToString(), "z");
