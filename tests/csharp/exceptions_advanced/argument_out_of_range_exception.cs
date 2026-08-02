// vybe-test: csharp/exceptions_advanced/argument_out_of_range_exception
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    throw new ArgumentOutOfRangeException("index", "too big");
} catch (ArgumentOutOfRangeException e) {
    __Check((e.ParamName).ToString(), "index");
}
