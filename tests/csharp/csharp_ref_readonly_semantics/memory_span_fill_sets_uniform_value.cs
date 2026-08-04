// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_fill_sets_uniform_value
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

var memory=new System.Memory<int>(new int[]{0,0,0}); memory.Span.Fill(4); __P((memory.Span[2]).ToString());
__Check("4");
