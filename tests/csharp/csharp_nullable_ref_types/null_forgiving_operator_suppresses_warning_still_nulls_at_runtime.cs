// vybe-test: csharp/csharp_nullable_ref_types/null_forgiving_operator_suppresses_warning_still_nulls_at_runtime
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
string r="ok";
try{__P((s!.Length).ToString());}
catch(System.NullReferenceException){r="null";}
__P((r).ToString());
__Check("null");
