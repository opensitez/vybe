// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_coalesce_uses_right_when_left_is_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? left = null;
__Check((left ?? 100).ToString(), "100");
