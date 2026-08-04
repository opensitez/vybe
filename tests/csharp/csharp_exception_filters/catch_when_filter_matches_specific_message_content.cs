// vybe-test: csharp/csharp_exception_filters/catch_when_filter_matches_specific_message_content
// origin: languages/csharp/tests/csharp/test_csharp_exception_filters.rs

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

try {
    throw new System.Exception("code=404");
} catch (System.Exception e) when (e.Message.Contains("404")) {
    __P(("not found").ToString());
} catch (System.Exception) {
    __P(("other").ToString());
}
__Check("not found");
