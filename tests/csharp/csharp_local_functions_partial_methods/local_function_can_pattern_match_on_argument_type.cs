// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_pattern_match_on_argument_type
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Describe(object value) { return value is int number ? (number * 2).ToString() : "other"; } __Check((Describe(6)).ToString(), "12");
