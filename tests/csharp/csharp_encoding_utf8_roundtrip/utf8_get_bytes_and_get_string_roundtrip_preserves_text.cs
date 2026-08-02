// vybe-test: csharp/csharp_encoding_utf8_roundtrip/utf8_get_bytes_and_get_string_roundtrip_preserves_text
// origin: languages/csharp/tests/csharp/test_csharp_encoding_utf8_roundtrip.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var encoding = System.Text.Encoding.UTF8;
var bytes = encoding.GetBytes("café");
var text = encoding.GetString(bytes);
__Check((text).ToString(), "café");
