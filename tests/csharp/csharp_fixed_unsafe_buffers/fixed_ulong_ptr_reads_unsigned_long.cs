// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_ulong_ptr_reads_unsigned_long
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

ulong[] arr={18446744073709551615UL,0UL}; unsafe{fixed(ulong* ptr=&arr[1]){__P((*ptr).ToString());}}
__Check("0");
