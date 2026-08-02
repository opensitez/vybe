// vybe-test: csharp/csharp_exception_filters/inner_exception_chains_to_outer_exception_cause
// origin: languages/csharp/tests/csharp/test_csharp_exception_filters.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    try {
        throw new System.Exception("root cause");
    } catch (System.Exception inner) {
        throw new System.InvalidOperationException("wrapped", inner);
    }
} catch (System.InvalidOperationException outer) {
    __Check((outer.Message).ToString(), "wrapped");
    __Check((outer.InnerException.Message).ToString(), "root cause");
}
