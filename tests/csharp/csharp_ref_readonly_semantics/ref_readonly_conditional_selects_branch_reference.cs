// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_conditional_selects_branch_reference
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

int[] arr={1,2,3}; bool pickSecond=true; ref readonly int chosen=ref (pickSecond?ref arr[1]:ref arr[0]); __P((chosen).ToString());
__Check("2");
