// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_distance_between_elements_is_one
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

int[] arr={1,2,3}; unsafe{fixed(int* a=&arr[0]){fixed(int* b=&arr[1]){__P((b-a).ToString());}}}
__Check("1");
