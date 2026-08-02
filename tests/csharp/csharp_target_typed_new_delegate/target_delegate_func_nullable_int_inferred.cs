// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_nullable_int_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int?, int> orZero = n => n ?? 0;
__Check((orZero(null)).ToString(), "0"); __Check((orZero(7)).ToString(), "7");
