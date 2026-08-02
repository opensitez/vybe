// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
string feature = "jagged_array_patterns"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
