// vybe-test: csharp/csharp_null_handling/null_conditional_invoke_on_event_is_safe
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action callback = null;
callback?.Invoke();
__Check(("safe").ToString(), "safe");
