// vybe-test: csharp/csharp_local_functions_partial_methods/static_local_function_does_not_capture_outer_state
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static int Triple(int value) { return value * 3; } __Check((Triple(4)).ToString(), "12");
