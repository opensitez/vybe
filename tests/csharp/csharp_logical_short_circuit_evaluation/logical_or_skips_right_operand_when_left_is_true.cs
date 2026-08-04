// vybe-test: csharp/csharp_logical_short_circuit_evaluation/logical_or_skips_right_operand_when_left_is_true
// origin: languages/csharp/tests/csharp/test_csharp_logical_short_circuit_evaluation.rs

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

int calls = 0;
bool Right() { calls++; return true; }
bool result = true || Right();
__P((result ? "T" : "F").ToString());
__P((calls).ToString());
__Check("T\n0");
