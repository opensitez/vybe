// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_chained_invocation
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int, int> inc = x => x + 1;
System.Func<int, int> twice = x => inc(inc(x));
__Check((twice(3)).ToString(), "5");
