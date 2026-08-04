// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_mismatch_reports_first_difference_index
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

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

System.ReadOnlySpan<int> a=stackalloc int[3]{1,2,3}; System.ReadOnlySpan<int> b=stackalloc int[3]{1,9,3}; __P((a.Mismatch(b)).ToString());
__Check("1");
