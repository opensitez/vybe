// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_be_recursive
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Fib(int n) { return n <= 1 ? n : Fib(n - 1) + Fib(n - 2); } __Check((Fib(6)).ToString(), "8");
