// vybe-test: csharp/csharp_char_operations/char_is_punctuation_true_for_dot
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((char.IsPunctuation('.')).ToString(), "True");
