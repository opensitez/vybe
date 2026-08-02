// vybe-test: csharp/csharp_delegate_types/delegate_null_check_before_invoke_prevents_null_reference
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action handler = null;
handler?.Invoke();
__Check(("safe").ToString(), "safe");
