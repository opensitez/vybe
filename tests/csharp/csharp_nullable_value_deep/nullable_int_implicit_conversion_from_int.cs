// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_implicit_conversion_from_int
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x=33; int? n=x; __Check((n).ToString(), "33");
