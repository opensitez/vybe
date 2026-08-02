// vybe-test: csharp/csharp_null_propagation/delegate_null_conditional_invocation_calls_delegate_when_present
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; Action action = () => __Check(("ran").ToString(), "ran"); action?.Invoke();
