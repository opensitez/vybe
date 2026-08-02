// vybe-test: csharp/csharp_lambda_loop_capture_semantics/lambda_mutating_captured_local_is_visible_to_later_invocations
// origin: languages/csharp/tests/csharp/test_csharp_lambda_loop_capture_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System;
int tally = 0;
Action bump = () => { tally++; };
bump();
bump();
__Check((tally).ToString(), "2");
