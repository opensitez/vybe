// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_starts_with_prefix_bytes
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes=u8"prefix"; __Check((bytes.StartsWith(u8"pre")).ToString(), "True");
