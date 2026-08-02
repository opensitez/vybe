// vybe-test: csharp/csharp_encoding/get_byte_count_reflects_character_byte_width
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = System.Text.Encoding.UTF8.GetByteCount("café");
__Check((n > 4).ToString(), "True");
