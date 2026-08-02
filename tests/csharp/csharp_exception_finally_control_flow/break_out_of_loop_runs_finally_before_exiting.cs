// vybe-test: csharp/csharp_exception_finally_control_flow/break_out_of_loop_runs_finally_before_exiting
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

string trace = "";
for (int i = 0; i < 3; i++) {
    try {
        trace += "body;";
        break;
    } finally {
        trace += "cleanup;";
    }
}
trace += "after";
Console.WriteLine(trace);
