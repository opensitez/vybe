// vybe-test: csharp/csharp_error_handling/try_catch_basic
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    throw new Exception("oops");
} catch (Exception e) {
    __Check((e.Message).ToString(), "oops");
}
