// vybe-test: csharp/csharp_nullable_semantics/null_coalescing_assign_only_sets_when_currently_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n = null; n ??= 5; __Check((n).ToString(), "5");
