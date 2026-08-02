// vybe-test: csharp/csharp_exception_finally_control_flow/exception_rethrown_from_catch_is_handled_by_enclosing_try
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
        throw new Exception("first");
    } catch (Exception) {
        trace += "inner;";
        throw new Exception("second");
    }
} catch (Exception e) {
    trace += "outer:" + e.Message;
}
__Check((trace).ToString(), "inner;outer:second");
