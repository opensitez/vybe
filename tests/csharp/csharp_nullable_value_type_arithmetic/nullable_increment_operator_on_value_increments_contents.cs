// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_increment_operator_on_value_increments_contents
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? value = 10;
value++;
__Check((value).ToString(), "11");
