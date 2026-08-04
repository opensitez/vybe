// vybe-test: csharp/csharp_value_ref_semantics/null_reference_throws_null_reference_exception
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

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

string r="";
try{string s=null;int len=s.Length;}
catch(System.NullReferenceException){r="null";}
__P((r).ToString());
__Check("null");
