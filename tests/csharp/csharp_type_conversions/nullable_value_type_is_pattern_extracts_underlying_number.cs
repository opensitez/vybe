// vybe-test: csharp/csharp_type_conversions/nullable_value_type_is_pattern_extracts_underlying_number
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? maybe = 30; if (maybe is int value) __Check((value / 3).ToString(), "10");
