// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_increment_operator_on_null_stays_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? value = null;
value++;
__Check((value.HasValue).ToString(), "False");
