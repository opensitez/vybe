// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
string feature = "string_builder_usage:20"; __Check((feature.Length >= 1).ToString(), "True");
