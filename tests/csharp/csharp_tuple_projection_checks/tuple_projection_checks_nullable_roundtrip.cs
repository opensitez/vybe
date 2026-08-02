// vybe-test: csharp/csharp_tuple_projection_checks/tuple_projection_checks_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_tuple_projection_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_projection_checks
int? maybe = 36; __Check((maybe.HasValue && maybe.Value == 36).ToString(), "True");
