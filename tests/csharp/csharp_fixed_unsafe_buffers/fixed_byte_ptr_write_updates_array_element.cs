// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_write_updates_array_element
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

byte[] arr={1,2,3}; unsafe{fixed(byte* ptr=&arr[0]){ptr[1]=99;}} __P((arr[1]).ToString());
__Check("99");
