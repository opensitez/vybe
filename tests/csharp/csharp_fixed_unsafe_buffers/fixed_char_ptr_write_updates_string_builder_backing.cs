// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_char_ptr_write_updates_string_builder_backing
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

char[] arr={'x','y'}; unsafe{fixed(char* ptr=&arr[0]){ptr[1]='z';}} __P((arr[1]).ToString());
__Check("122");
