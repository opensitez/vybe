// vybe-test: csharp/csharp_local_functions_partial_methods/static_local_function_is_callable_without_capturing_outer_state
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Compute() { static int Double(int n) => n * 2; return Double(4); } __Check((Compute()).ToString(), "8");
