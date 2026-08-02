// vybe-test: csharp/csharp_convert_uri_path/encoding_utf8_roundtrips_unicode_text
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes = System.Text.Encoding.UTF8.GetBytes("café");
var text = System.Text.Encoding.UTF8.GetString(bytes);
__Check((text).ToString(), "café");
