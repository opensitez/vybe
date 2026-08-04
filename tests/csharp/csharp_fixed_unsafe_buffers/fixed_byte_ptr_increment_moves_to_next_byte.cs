// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_increment_moves_to_next_byte
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

byte[] arr={7,8,9}; unsafe{fixed(byte* ptr=&arr[0]){ptr++; __P((*ptr).ToString());}}
__Check("8");
