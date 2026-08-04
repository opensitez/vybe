// vybe-test: csharp/csharp_exceptions_hierarchy/aggregate_exception_wraps_multiple_inner_exceptions
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var ae=new System.AggregateException(
    new System.Exception("one"),
    new System.Exception("two"));
__P((ae.InnerExceptions.Count).ToString());
__Check("2");
