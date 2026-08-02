// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
string feature = "jagged_array_patterns"; __Check((feature[0] == feature[0]).ToString(), "True");
