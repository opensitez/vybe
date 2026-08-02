// vybe-test: csharp/csharp_delegate_variance/action_contravariant_stored_in_base_reference
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> baseAct=v=>__Check((v).ToString(), "hold"); System.Action<string> derivedAct=baseAct; object holder=derivedAct; ((System.Action<string>)holder)("hold");
