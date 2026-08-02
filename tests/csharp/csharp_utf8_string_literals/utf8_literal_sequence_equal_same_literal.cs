// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_sequence_equal_same_literal
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=u8"same"; var b=u8"same"; __Check((a.SequenceEqual(b)).ToString(), "True");
