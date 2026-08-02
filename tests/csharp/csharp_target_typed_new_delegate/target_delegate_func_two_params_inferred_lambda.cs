// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_two_params_inferred_lambda
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int, int, int> mul = (a, b) => a * b;
__Check((mul(3, 5)).ToString(), "15");
