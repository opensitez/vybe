// vybe-test: csharp/csharp_delegate_variance/action_contravariant_uppercase_via_object
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> w=v=>__Check((((string)v).ToUpper()).ToString(), "HI"); System.Action<string> n=w; n("hi");
