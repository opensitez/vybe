// vybe-test: csharp/csharp_exception_filters/rethrow_preserves_original_stack_trace
// origin: languages/csharp/tests/csharp/test_csharp_exception_filters.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "";
try {
    try {
        throw new System.Exception("original");
    } catch (System.Exception) {
        throw;
    }
} catch (System.Exception e) {
    result = e.Message;
}
__Check((result).ToString(), "original");
