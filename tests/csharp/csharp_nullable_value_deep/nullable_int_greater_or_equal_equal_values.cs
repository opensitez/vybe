// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_greater_or_equal_equal_values
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? a=4; int? b=4; __Check((a>=b).ToString(), "True");
