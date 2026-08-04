// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_extension_method_on_span
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

static class SpanExt{public static int First(ref readonly System.Span<int> span)=>span.Length>0?span[0]:-1;} System.Span<int> s=stackalloc int[2]{5,6}; __P((SpanExt.First(ref s)).ToString());
__Check("5");
