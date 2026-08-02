// vybe-test: csharp/csharp_encoding/utf8_roundtrip_string_through_bytes_and_back
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes = System.Text.Encoding.UTF8.GetBytes("hello");
__Check((System.Text.Encoding.UTF8.GetString(bytes)).ToString(), "hello");
