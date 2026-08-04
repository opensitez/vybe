// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_two_fixed_blocks_same_array
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

byte[] arr={1,2,3}; unsafe{fixed(byte* a=&arr[0]){fixed(byte* b=&arr[1]){__P((*a+b[0]).ToString());}}}
__Check("3");
