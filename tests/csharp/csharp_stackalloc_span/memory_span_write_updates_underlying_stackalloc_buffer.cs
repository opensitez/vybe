// vybe-test: csharp/csharp_stackalloc_span/memory_span_write_updates_underlying_stackalloc_buffer
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

System.Memory<int> mem=new System.Memory<int>(stackalloc int[2]{1,2}); mem.Span[0]=77; __P((mem.Span[0]).ToString());
__Check("77");
