// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_swap_two_elements
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

int[] arr={1,9}; unsafe{fixed(int* ptr=&arr[0]){int t=ptr[0]; ptr[0]=ptr[1]; ptr[1]=t;}} __P((arr[0]).ToString()); __P((arr[1]).ToString());
__Check("9\n1");
