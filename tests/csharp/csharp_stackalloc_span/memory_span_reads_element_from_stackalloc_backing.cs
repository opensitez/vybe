// vybe-test: csharp/csharp_stackalloc_span/memory_span_reads_element_from_stackalloc_backing
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

System.Memory<int> mem=new System.Memory<int>(stackalloc int[3]{4,5,6}); __P((mem.Span[1]).ToString());
__Check("5");
