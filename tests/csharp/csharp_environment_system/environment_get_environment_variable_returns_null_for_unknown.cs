// vybe-test: csharp/csharp_environment_system/environment_get_environment_variable_returns_null_for_unknown
// origin: languages/csharp/tests/csharp/test_csharp_environment_system.rs

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

var v=System.Environment.GetEnvironmentVariable("__VYBE_NOSUCH_VAR__123");
__P((v==null).ToString());
__Check("True");
