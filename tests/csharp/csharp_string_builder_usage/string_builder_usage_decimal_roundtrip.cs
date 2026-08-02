// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
