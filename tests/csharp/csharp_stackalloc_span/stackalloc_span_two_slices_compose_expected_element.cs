// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_two_slices_compose_expected_element
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

System.Span<int> span=stackalloc int[6]{1,2,3,4,5,6}; var inner=span.Slice(2,2); __P((inner[1]).ToString());
__Check("4");
