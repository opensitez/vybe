// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_returning_delegate_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int, System.Func<int, int>> scale = factor => n => n * factor;
var triple = scale(3);
__Check((triple(4)).ToString(), "12");
