// vybe-test: csharp/csharp_nullable_semantics/null_coalescing_operator_returns_right_side_when_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n = null; __Check((n ?? 99).ToString(), "99");
