// vybe-test: csharp/csharp_ref_return_semantics/ref_return_chains_to_second_ref_local_without_copying_value
// origin: languages/csharp/tests/csharp/test_csharp_ref_return_semantics.rs

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

int[] values = { 10, 20 };
ref int First() => ref values[0];
ref int alias = ref First();
alias = 99;
__P((values[0]).ToString());
__Check("99");
