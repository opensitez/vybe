// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_have_ref_parameter
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Increment(ref int value) { value++; } int count = 7; Increment(ref count); __Check((count).ToString(), "8");
