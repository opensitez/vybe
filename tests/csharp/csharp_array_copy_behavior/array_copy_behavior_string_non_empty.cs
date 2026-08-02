// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_copy_behavior
string feature = "array_copy_behavior"; __Check((feature.Length > 0).ToString(), "True");
