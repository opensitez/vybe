// vybe-test: csharp/csharp_throw_in_null_coalescing/null_coalescing_throw_expression_skipped_when_left_has_value
// origin: languages/csharp/tests/csharp/test_csharp_throw_in_null_coalescing.rs

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

string? present = "ok";
string value = present ?? throw new System.Exception("fail");
__P((value).ToString());
__Check("ok");
