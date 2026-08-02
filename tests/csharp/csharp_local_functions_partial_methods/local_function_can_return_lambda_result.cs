// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_return_lambda_result
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Compute() { System.Func<int> read = () => 9; return read() + 1; } __Check((Compute()).ToString(), "10");
