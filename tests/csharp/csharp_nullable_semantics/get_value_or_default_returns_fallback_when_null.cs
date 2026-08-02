// vybe-test: csharp/csharp_nullable_semantics/get_value_or_default_returns_fallback_when_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n = null; __Check((n.GetValueOrDefault(-1)).ToString(), "-1");
