// vybe-test: csharp/csharp_char_operations/char_to_upper_converts_lowercase
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((char.ToUpper('b')).ToString(), "B");
