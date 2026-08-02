// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_get_value_or_default_returns_zero_for_null_int
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? value = null;
__Check((value.GetValueOrDefault()).ToString(), "0");
__Check((value.GetValueOrDefault(99)).ToString(), "99");
