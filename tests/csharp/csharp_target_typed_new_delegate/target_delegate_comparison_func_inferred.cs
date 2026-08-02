// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_comparison_func_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int, int, bool> less = (a, b) => a < b;
__Check((less(2, 5)).ToString(), "True"); __Check((less(9, 1)).ToString(), "False");
