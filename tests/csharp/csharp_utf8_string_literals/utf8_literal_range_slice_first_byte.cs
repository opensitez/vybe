// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_range_slice_first_byte
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes=u8"abcdef"; var slice=bytes[2..5]; __Check((slice[0]).ToString(), "99");
