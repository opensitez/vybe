// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_has_value_true_for_zero
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n=0; __Check((n.HasValue).ToString(), "True");
