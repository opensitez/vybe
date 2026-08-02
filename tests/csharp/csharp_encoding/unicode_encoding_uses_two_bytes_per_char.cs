// vybe-test: csharp/csharp_encoding/unicode_encoding_uses_two_bytes_per_char
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes = System.Text.Encoding.Unicode.GetBytes("A");
__Check((bytes.Length).ToString(), "2");
