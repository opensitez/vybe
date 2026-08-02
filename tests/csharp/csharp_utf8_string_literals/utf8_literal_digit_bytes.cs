// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_digit_bytes
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes=u8"123"; __Check((bytes[0]).ToString(), "49"); __Check((bytes[2]).ToString(), "51");
