// vybe-test: csharp/csharp_span_indexing/memory_slice_reads_correct_element_via_span
// origin: languages/csharp/tests/csharp/test_csharp_span_indexing.rs

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

var memory = new System.Memory<int>(new int[] { 5, 6, 7 });
__P((memory.Span[2]).ToString());
__Check("7");
