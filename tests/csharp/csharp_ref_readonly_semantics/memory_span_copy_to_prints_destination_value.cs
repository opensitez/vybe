// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_copy_to_prints_destination_value
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

var src=new System.Memory<int>(new int[]{7,8}); int[] dst=new int[2]; src.Span.CopyTo(dst); __P((dst[1]).ToString());
__Check("8");
