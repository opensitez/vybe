// vybe-test: csharp/csharp_delegate_variance/action_contravariant_double_to_object
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> log=v=>__Check(((double)v==3.14).ToString(), "True"); System.Action<double> logD=log; logD(3.14);
