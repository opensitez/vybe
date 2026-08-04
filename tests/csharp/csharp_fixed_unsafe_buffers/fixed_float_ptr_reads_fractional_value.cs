// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_float_ptr_reads_fractional_value
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

float[] arr={1.5f,2.5f}; unsafe{fixed(float* ptr=&arr[0]){__P((*ptr==1.5f).ToString());}}
__Check("True");
