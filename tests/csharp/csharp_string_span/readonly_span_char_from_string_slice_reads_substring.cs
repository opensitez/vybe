// vybe-test: csharp/csharp_string_span/readonly_span_char_from_string_slice_reads_substring
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

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

string s="hello world";
System.ReadOnlySpan<char> span=s.AsSpan(6,5);
__P((span.ToString()).ToString());
__Check("world");
