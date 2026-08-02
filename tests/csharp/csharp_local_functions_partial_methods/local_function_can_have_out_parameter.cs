// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_have_out_parameter
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Split(int value, out int left, out int right) { left = value / 2; right = value - left; } Split(9, out var left, out var right); __Check((left).ToString(), "4"); __Check((right).ToString(), "5");
