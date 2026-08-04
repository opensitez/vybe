// vybe-test: csharp/csharp_exception_finally_control_flow/catch_filter_when_clause_skips_handler_for_non_matching_predicate
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

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

string label = "start";
try {
    throw new Exception("code-404");
} catch (Exception e) when (e.Message.Contains("500")) {
    label = "wrong";
} catch (Exception e) when (e.Message.Contains("404")) {
    label = "matched";
}
__P((label).ToString());
__Check("matched");
