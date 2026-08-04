// vybe-test: csharp/csharp_stackalloc_span/span_slice_start_and_length_reads_subrange
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

System.Span<int> span=stackalloc int[5]{10,20,30,40,50}; var mid=span.Slice(1,2); __P((mid[0]).ToString()); __P((mid[1]).ToString());
__Check("20\n30");
