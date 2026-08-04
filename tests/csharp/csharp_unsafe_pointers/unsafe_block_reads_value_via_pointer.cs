// vybe-test: csharp/csharp_unsafe_pointers/unsafe_block_reads_value_via_pointer
// origin: languages/csharp/tests/csharp/test_csharp_unsafe_pointers.rs

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

unsafe{
    int x=42;
    int* p=&x;
    __P((*p).ToString());
}
__Check("42");
