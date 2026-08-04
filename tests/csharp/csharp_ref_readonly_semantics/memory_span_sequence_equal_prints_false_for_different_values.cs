// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_sequence_equal_prints_false_for_different_values
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

var left=new System.Memory<int>(new int[]{1,2}); var right=new System.Memory<int>(new int[]{1,9}); __P((left.Span.SequenceEqual(right.Span)).ToString());
__Check("False");
