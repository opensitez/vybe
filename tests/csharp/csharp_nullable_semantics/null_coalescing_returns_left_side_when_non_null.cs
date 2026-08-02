// vybe-test: csharp/csharp_nullable_semantics/null_coalescing_returns_left_side_when_non_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n = 7; __Check((n ?? 99).ToString(), "7");
