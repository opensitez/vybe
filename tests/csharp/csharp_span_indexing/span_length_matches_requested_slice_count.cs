// vybe-test: csharp/csharp_span_indexing/span_length_matches_requested_slice_count
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

int[] data = { 1, 2, 3, 4 };
var span = data.AsSpan(1, 2);
__P((span.Length).ToString());
__Check("2");
