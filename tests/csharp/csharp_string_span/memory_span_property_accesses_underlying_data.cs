// vybe-test: csharp/csharp_string_span/memory_span_property_accesses_underlying_data
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

System.Memory<int> m=new int[]{7,8,9};
__P((m.Span[1]).ToString());
__Check("8");
