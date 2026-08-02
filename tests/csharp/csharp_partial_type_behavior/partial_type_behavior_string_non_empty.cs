// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
string feature = "partial_type_behavior"; __Check((feature.Length > 0).ToString(), "True");
