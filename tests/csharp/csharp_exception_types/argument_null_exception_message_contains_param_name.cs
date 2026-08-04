// vybe-test: csharp/csharp_exception_types/argument_null_exception_message_contains_param_name
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

try { throw new System.ArgumentNullException("value"); }
catch(System.ArgumentNullException e) { __P((e.ParamName).ToString()); }
__Check("value");
