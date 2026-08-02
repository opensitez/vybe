// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_use_default_parameter_value
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Add(int left, int right = 10) { return left + right; } __Check((Add(5)).ToString(), "15");
