// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
string feature = "string_builder_usage"; __Check((feature[0] == feature[0]).ToString(), "True");
