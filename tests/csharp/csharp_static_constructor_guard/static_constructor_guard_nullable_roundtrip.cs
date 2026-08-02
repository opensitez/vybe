// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
int? maybe = 69; __Check((maybe.HasValue && maybe.Value == 69).ToString(), "True");
