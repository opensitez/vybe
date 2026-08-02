// vybe-test: csharp/csharp_nullable_semantics/nullable_value_type_cast_to_non_nullable_succeeds_with_value
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n = 10; int x = (int)n; __Check((x).ToString(), "10");
