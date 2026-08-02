// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_return_tuple_result
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

(int, int) Pair() { return (2, 5); } var result = Pair(); __Check((result.Item1 + result.Item2).ToString(), "7");
