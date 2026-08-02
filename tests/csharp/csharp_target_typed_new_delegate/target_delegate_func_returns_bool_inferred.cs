// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_returns_bool_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int, bool> positive = n => n > 0;
__Check((positive(1)).ToString(), "True"); __Check((positive(-1)).ToString(), "False");
