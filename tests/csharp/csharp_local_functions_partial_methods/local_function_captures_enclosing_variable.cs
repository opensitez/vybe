// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_captures_enclosing_variable
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int factor = 3; int Scale(int value) { return value * factor; } __Check((Scale(5)).ToString(), "15");
