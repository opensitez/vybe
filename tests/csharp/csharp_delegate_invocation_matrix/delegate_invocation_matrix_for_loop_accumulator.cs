// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

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

// delegate_invocation_matrix
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } __P((sum == 3).ToString());
__Check("True");
