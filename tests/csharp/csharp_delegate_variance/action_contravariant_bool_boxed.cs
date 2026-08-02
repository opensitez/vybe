// vybe-test: csharp/csharp_delegate_variance/action_contravariant_bool_boxed
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> log=v=>__Check((v).ToString(), "True"); System.Action<bool> logB=log; logB(true);
