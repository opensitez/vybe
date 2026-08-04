// vybe-test: csharp/csharp_throw_in_null_coalescing/chained_null_coalescing_throw_only_evaluates_when_all_prior_operands_null
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

string? a = null;
string? b = null;
try {
    string value = a ?? b ?? throw new System.Exception("both-null");
    __P((value).ToString());
} catch (System.Exception) {
    __P(("caught").ToString());
}
__Check("caught");
