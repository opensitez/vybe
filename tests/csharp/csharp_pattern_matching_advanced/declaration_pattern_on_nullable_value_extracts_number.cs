// vybe-test: csharp/csharp_pattern_matching_advanced/declaration_pattern_on_nullable_value_extracts_number
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? value = 7; if (value is int number) __Check((number + 1).ToString(), "8");
