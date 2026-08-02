// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_decodes_to_string_via_encoding
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes=u8"hello"; __Check((System.Text.Encoding.UTF8.GetString(bytes)).ToString(), "hello");
