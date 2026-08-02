// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
int seed = 20; int right = seed + 1; __Check((seed < right).ToString(), "True");
