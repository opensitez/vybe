// vybe-test: csharp/csharp_unsafe_pointers/fixed_statement_pins_array_for_pointer_arithmetic
// origin: languages/csharp/tests/csharp/test_csharp_unsafe_pointers.rs

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

int[] arr={10,20,30};
unsafe{
    fixed(int* p=arr){
        __P((*(p+1)).ToString());
    }
}
__Check("20");
