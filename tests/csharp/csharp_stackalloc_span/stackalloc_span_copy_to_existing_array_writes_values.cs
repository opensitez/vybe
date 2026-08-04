// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_copy_to_existing_array_writes_values
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

System.Span<int> src=stackalloc int[2]{8,9}; int[] dst=new int[2]; src.CopyTo(dst); __P((dst[1]).ToString());
__Check("9");
