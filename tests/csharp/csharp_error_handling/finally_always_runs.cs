// vybe-test: csharp/csharp_error_handling/finally_always_runs
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    __Check(("before").ToString(), "before");
    throw new Exception("err");
} catch (Exception e) {
    __Check(("caught").ToString(), "caught");
} finally {
    __Check(("always").ToString(), "always");
}
