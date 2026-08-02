// vybe-test: csharp/csharp_nullable_semantics/nullable_int_has_value_true_when_assigned
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n = 5; __Check((n.HasValue).ToString(), "True");
