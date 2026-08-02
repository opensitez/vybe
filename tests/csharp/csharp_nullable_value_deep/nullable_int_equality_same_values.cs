// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_equality_same_values
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? a=5; int? b=5; __Check((a==b).ToString(), "True");
