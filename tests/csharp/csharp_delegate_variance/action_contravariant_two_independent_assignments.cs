// vybe-test: csharp/csharp_delegate_variance/action_contravariant_two_independent_assignments
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> a=v=>__Check(("a").ToString(), "a"); System.Action<object> b=v=>__Check(("b").ToString(), "b"); System.Action<string> sa=a; System.Action<string> sb=b; sa(""); sb("");
