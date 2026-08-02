// vybe-test: csharp/csharp_lambda_loop_capture_semantics/lambda_passed_to_method_receives_current_binding_not_snapshot_at_creation
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System;
int total = 1;
Action add = () => total += 4;
void Apply(Action work) { work(); }
Apply(add);
__Check((total).ToString(), "5");
