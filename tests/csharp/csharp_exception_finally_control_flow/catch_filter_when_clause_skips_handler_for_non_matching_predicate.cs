// vybe-test: csharp/csharp_exception_finally_control_flow/catch_filter_when_clause_skips_handler_for_non_matching_predicate
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string label = "start";
try {
    throw new Exception("code-404");
} catch (Exception e) when (e.Message.Contains("500")) {
    label = "wrong";
} catch (Exception e) when (e.Message.Contains("404")) {
    label = "matched";
}
__Check((label).ToString(), "matched");
