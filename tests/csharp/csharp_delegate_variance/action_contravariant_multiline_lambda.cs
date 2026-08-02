// vybe-test: csharp/csharp_delegate_variance/action_contravariant_multiline_lambda
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> wide=v=>{__Check((v).ToString(), "99");}; System.Action<int> narrow=wide; narrow(99);
