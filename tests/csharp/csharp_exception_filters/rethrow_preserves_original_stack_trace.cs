// vybe-test: csharp/csharp_exception_filters/rethrow_preserves_original_stack_trace
// origin: languages/csharp/tests/csharp/test_csharp_exception_filters.rs

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
__P((result).ToString());
__Check("original");
