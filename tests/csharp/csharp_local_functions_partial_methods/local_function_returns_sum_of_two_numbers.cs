// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_returns_sum_of_two_numbers
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Add(int left, int right) { return left + right; } __Check((Add(3, 4)).ToString(), "7");
