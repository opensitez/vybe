// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_write_to_outer_variable
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 0; void Add(int value) { total += value; } Add(4); Add(6); __Check((total).ToString(), "10");
