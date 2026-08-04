// vybe-test: csharp/csharp_exception_types/exception_message_survives_catch_and_rethrow_as_inner
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

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

string msg = "";
try {
    try { throw new System.Exception("root"); }
    catch(System.Exception e) { throw new System.Exception("wrap", e); }
} catch(System.Exception outer) { msg = outer.InnerException.Message; }
__P((msg).ToString());
__Check("root");
