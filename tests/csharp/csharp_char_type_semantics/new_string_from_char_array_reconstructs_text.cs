// vybe-test: csharp/csharp_char_type_semantics/new_string_from_char_array_reconstructs_text
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char[] data = { 'h', 'i' };
__Check((new string(data)).ToString(), "hi");
