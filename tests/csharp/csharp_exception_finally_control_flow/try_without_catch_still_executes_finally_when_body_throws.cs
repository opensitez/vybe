// vybe-test: csharp/csharp_exception_finally_control_flow/try_without_catch_still_executes_finally_when_body_throws
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string trace = "";
try {
    try {
        throw new Exception("fail");
    } finally {
        trace += "finally;";
    }
} catch (Exception) {
    trace += "handled;";
}
__Check((trace).ToString(), "finally;handled;");
