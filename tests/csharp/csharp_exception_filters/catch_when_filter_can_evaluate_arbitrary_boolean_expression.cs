// vybe-test: csharp/csharp_exception_filters/catch_when_filter_can_evaluate_arbitrary_boolean_expression
// origin: languages/csharp/tests/csharp/test_csharp_exception_filters.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int threshold = 10;
try {
    throw new System.InvalidOperationException("value=15");
} catch (System.InvalidOperationException e) when (threshold < 20) {
    __Check(("caught with threshold").ToString(), "caught with threshold");
}
