// vybe-test: csharp/csharp_ref_out_in/ref_local_aliases_array_element
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

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

int[] arr={1,2,3};
ref int second=ref arr[1];
second=99;
__P((arr[1]).ToString());
__Check("99");
