// vybe-test: csharp/csharp_type_conversions/integer_to_char_cast_produces_character
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 66; __Check(((char)value).ToString(), "B");
