// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_with_expression_body_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Double(int value) => value * 2; __Check((Double(9)).ToString(), "18");
