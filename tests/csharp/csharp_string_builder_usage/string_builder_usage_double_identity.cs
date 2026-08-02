// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
double seed = 20; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
