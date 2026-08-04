// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_compare_elements_via_pointers
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

int[] arr={5,5,9}; unsafe{fixed(int* ptr=&arr[0]){__P((ptr[0]==ptr[1]).ToString()); __P((ptr[0]==ptr[2]).ToString());}}
__Check("True\nFalse");
