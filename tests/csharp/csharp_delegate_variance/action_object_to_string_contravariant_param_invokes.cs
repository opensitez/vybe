// vybe-test: csharp/csharp_delegate_variance/action_object_to_string_contravariant_param_invokes
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<object> logObject=v=>__Check((v).ToString(), "typed"); System.Action<string> logString=logObject; logString("typed");
