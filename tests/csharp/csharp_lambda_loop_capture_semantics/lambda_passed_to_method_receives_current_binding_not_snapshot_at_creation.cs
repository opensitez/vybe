// vybe-test: csharp/csharp_lambda_loop_capture_semantics/lambda_passed_to_method_receives_current_binding_not_snapshot_at_creation
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

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

using System;
int total = 1;
Action add = () => total += 4;
void Apply(Action work) { work(); }
Apply(add);
__P((total).ToString());
__Check("5");
