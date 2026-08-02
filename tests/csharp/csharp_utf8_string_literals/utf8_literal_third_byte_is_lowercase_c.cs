// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_third_byte_is_lowercase_c
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes=u8"abc"; __Check((bytes[2]).ToString(), "99");
