// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_inferred_from_lambda_assignment
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int, int> triple = x => x * 3;
__Check((triple(4)).ToString(), "12");
