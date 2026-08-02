// vybe-test: csharp/csharp_exception_finally_control_flow/finally_that_throws_during_return_propagates_past_the_try
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void M() {
    try {
        try {
            __Check(("body").ToString(), "body");
            return;
        } finally {
            __Check(("finally").ToString(), "finally");
            throw new Exception("boom");
        }
    } catch (Exception) {
        __Check(("caught").ToString(), "caught");
    }
}
M();
