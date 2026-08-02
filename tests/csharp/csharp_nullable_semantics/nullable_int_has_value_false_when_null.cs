// vybe-test: csharp/csharp_nullable_semantics/nullable_int_has_value_false_when_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n = null; __Check((n.HasValue).ToString(), "False");
