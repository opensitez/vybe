// vybe-test: csharp/csharp_logical_short_circuit_evaluation/and_short_circuits_before_or_evaluates_fallback_operand
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

int trace = 0;
bool A() { trace++; return false; }
bool B() { trace++; return true; }
bool C() { trace++; return true; }
bool value = A() && B() || C();
__P((value ? "T" : "F").ToString());
__P((trace).ToString());
__Check("T\n2");
