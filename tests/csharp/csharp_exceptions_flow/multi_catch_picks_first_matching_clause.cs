// vybe-test: csharp/csharp_exceptions_flow/multi_catch_picks_first_matching_clause
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_flow.rs

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
try{throw new System.ArgumentNullException("x");}
catch(System.ArgumentOutOfRangeException){r="range";}
catch(System.ArgumentNullException){r="null";}
catch(System.Exception){r="general";}
__P((r).ToString());
__Check("null");
