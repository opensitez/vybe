// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_address_of_element_equals_base_plus_offset
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

int[] arr={10,20,30}; unsafe{fixed(int* basePtr=&arr[0]){fixed(int* off=&arr[2]){__P((off-basePtr).ToString());}}}
__Check("2");
