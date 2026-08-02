// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_concatenation_not_supported_use_separate
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left=u8"ab"; var right=u8"cd"; __Check((left.Length+right.Length).ToString(), "4");
