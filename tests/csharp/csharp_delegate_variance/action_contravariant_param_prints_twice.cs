// vybe-test: csharp/csharp_delegate_variance/action_contravariant_param_prints_twice
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> once=v=>{__Check((v).ToString(), "x"); __Check((v).ToString(), "x");}; System.Action<string> twice=once; twice("x");
