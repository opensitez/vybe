// vybe-test: csharp/csharp_exceptions_hierarchy/aggregate_exception_wraps_multiple_inner_exceptions
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ae=new System.AggregateException(
    new System.Exception("one"),
    new System.Exception("two"));
__Check((ae.InnerExceptions.Count).ToString(), "2");
