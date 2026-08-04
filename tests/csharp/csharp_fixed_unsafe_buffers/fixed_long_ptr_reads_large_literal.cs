// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_long_ptr_reads_large_literal
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

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

long[] arr={10000000000L,2L}; unsafe{fixed(long* ptr=&arr[0]){__P((*ptr>0).ToString());}}
__Check("True");
