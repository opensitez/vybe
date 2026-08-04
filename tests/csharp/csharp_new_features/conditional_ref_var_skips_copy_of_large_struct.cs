// vybe-test: csharp/csharp_new_features/conditional_ref_var_skips_copy_of_large_struct
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

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

int[] arr = {1,2,3};
ref int val = ref arr[1];
val = 99;
__P((arr[1]).ToString());
__Check("99");
