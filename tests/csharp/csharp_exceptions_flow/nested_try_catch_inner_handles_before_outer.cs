// vybe-test: csharp/csharp_exceptions_flow/nested_try_catch_inner_handles_before_outer
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
try{
    try{throw new System.Exception("inner");}
    catch(System.Exception ex){r="inner:"+ex.Message; throw new System.Exception("outer");}
}catch(System.Exception ex){r+=" outer:"+ex.Message;}
__P((r).ToString());
__Check("inner:inner outer:outer");
