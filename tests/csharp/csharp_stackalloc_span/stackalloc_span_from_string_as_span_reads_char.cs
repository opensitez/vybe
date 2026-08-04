// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_from_string_as_span_reads_char
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

System.ReadOnlySpan<char> chars="abcd".AsSpan(1,2); __P((chars[0]).ToString()); __P((chars[1]).ToString());
__Check("b\nc");
