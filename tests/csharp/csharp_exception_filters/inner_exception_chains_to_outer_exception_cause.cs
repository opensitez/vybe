// vybe-test: csharp/csharp_exception_filters/inner_exception_chains_to_outer_exception_cause
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

try {
    try {
        throw new System.Exception("root cause");
    } catch (System.Exception inner) {
        throw new System.InvalidOperationException("wrapped", inner);
    }
} catch (System.InvalidOperationException outer) {
    __P((outer.Message).ToString());
    __P((outer.InnerException.Message).ToString());
}
__Check("wrapped\nroot cause");
