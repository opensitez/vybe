// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_equals_manual_byte_array
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes=u8"hi"; __Check((bytes[0]==104 && bytes[1]==105).ToString(), "True");
