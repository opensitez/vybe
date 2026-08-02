// vybe-test: csharp/csharp_encoding/utf8_get_bytes_returns_byte_array_for_ascii_text
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes = System.Text.Encoding.UTF8.GetBytes("hi");
__Check((bytes.Length).ToString(), "2");
