// vybe-test: csharp/csharp_logical_short_circuit_evaluation/logical_or_evaluates_right_operand_when_left_is_false
// origin: languages/csharp/tests/csharp/test_csharp_logical_short_circuit_evaluation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int calls = 0;
bool Right() { calls++; return true; }
bool result = false || Right();
__Check((result ? "T" : "F").ToString(), "T");
__Check((calls).ToString(), "1");
