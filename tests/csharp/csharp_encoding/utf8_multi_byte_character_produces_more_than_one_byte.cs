// vybe-test: csharp/csharp_encoding/utf8_multi_byte_character_produces_more_than_one_byte
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes = System.Text.Encoding.UTF8.GetBytes("€");
__Check((bytes.Length > 1).ToString(), "True");
