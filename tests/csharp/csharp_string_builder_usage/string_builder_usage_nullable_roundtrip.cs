// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
int? maybe = 20; __Check((maybe.HasValue && maybe.Value == 20).ToString(), "True");
