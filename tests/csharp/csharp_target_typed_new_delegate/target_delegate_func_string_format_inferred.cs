// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_string_format_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string, string, string> join = (a, b) => a + "-" + b;
__Check((join("x", "y")).ToString(), "x-y");
