// vybe-test: csharp/csharp_local_functions_partial_methods/generic_local_function_returns_typed_argument
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Echo<T>(T value) { return value; } __Check((Echo("generic")).ToString(), "generic");
