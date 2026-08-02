// vybe-test: csharp/csharp_char_operations/char_is_whitespace_true_for_space_and_tab
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((char.IsWhiteSpace(' ')).ToString(), "True"); __Check((char.IsWhiteSpace('\t')).ToString(), "True");
