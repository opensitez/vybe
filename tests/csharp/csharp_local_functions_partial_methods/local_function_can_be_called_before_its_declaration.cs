// vybe-test: csharp/csharp_local_functions_partial_methods/local_function_can_be_called_before_its_declaration
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Read()).ToString(), "ok"); string Read() { return "ok"; }
