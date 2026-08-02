// vybe-test: csharp/linq_lambdas/delegate_multicast
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Action<string> logger = msg => __Check(("LOG: " + msg).ToString(), "LOG: hello");
Action<string> printer = msg => __Check(("PRINT: " + msg).ToString(), "PRINT: hello");
Action<string> both = logger + printer;
both("hello");
