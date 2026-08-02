// vybe-test: csharp/csharp_null_propagation/delegate_null_conditional_invocation_skips_when_delegate_is_null
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; Action action = null; action?.Invoke(); __Check(("done").ToString(), "done");
