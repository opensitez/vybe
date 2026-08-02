// vybe-test: csharp/csharp_error_handling/nested_try_catch
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    try {
        throw new Exception("inner");
    } catch (Exception e) {
        __Check(("inner: " + e.Message).ToString(), "inner: inner");
        throw new Exception("rethrown");
    }
} catch (Exception e) {
    __Check(("outer: " + e.Message).ToString(), "outer: rethrown");
}
