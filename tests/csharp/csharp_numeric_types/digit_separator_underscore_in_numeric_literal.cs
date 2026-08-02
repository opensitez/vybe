// vybe-test: csharp/csharp_numeric_types/digit_separator_underscore_in_numeric_literal
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int million = 1_000_000; __Check((million).ToString(), "1000000");
