// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_byte_ptr_from_field_backed_array
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

class Holder{public byte[] Data={9,8};} var h=new Holder(); unsafe{fixed(byte* ptr=&h.Data[0]){__P((*ptr).ToString());}}
__Check("9");
