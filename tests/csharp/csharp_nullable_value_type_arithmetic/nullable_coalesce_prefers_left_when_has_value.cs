// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_coalesce_prefers_left_when_has_value
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? left = 8;
__Check((left ?? 100).ToString(), "8");
