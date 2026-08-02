// vybe-test: csharp/csharp_delegate_variance/action_contravariant_with_local_function
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Show(object o)=>__Check((o).ToString(), "fn"); System.Action<object> baseAct=Show; System.Action<string> derivedAct=baseAct; derivedAct("fn");
