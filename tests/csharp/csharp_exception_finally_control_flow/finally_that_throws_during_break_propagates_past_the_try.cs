// vybe-test: csharp/csharp_exception_finally_control_flow/finally_that_throws_during_break_propagates_past_the_try
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

string trace = "";
try {
    for (int i = 0; i < 3; i++) {
        try {
            trace += "body;";
            break;
        } finally {
            trace += "finally;";
            throw new Exception("boom");
        }
    }
    trace += "unreachable;";
} catch (Exception) {
    trace += "caught";
}
Console.WriteLine(trace);
