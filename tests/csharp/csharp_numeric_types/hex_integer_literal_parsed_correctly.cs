// vybe-test: csharp/csharp_numeric_types/hex_integer_literal_parsed_correctly
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 0xFF; __Check((n).ToString(), "255");
