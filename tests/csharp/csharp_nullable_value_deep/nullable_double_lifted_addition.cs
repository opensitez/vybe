// vybe-test: csharp/csharp_nullable_value_deep/nullable_double_lifted_addition
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double? a=1.5; double? b=2.5; __Check((a+b).ToString(), "4");
