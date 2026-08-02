// vybe-test: csharp/csharp_logical_short_circuit_evaluation/and_short_circuits_before_or_evaluates_fallback_operand
// origin: languages/csharp/tests/csharp/test_csharp_logical_short_circuit_evaluation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int trace = 0;
bool A() { trace++; return false; }
bool B() { trace++; return true; }
bool C() { trace++; return true; }
bool value = A() && B() || C();
__Check((value ? "T" : "F").ToString(), "T");
__Check((trace).ToString(), "2");
