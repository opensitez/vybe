// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
int? maybe = null; int fallback = maybe ?? 20; __Check((fallback == 20).ToString(), "True");
