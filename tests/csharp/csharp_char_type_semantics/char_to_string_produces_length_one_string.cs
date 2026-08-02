// vybe-test: csharp/csharp_char_type_semantics/char_to_string_produces_length_one_string
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char ch = 'x'; __Check((ch.ToString().Length).ToString(), "1");
