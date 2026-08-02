// vybe-test: csharp/csharp_char_operations/char_is_upper_distinguishes_case
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((char.IsUpper('A')).ToString(), "True"); __Check((char.IsUpper('a')).ToString(), "False");
