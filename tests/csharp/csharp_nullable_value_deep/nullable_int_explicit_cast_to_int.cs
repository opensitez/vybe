// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_explicit_cast_to_int
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n=12; int x=(int)n; __Check((x).ToString(), "12");
