// vybe-test: csharp/csharp_nullable_ref_types/nullable_reference_getvalueorlength_via_null_coalescing
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

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

string? s=null;
int len=s?.Length??-1;
__P((len).ToString());
__Check("-1");
