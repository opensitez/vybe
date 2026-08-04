// vybe-test: csharp/csharp_delegate_variance/action_contravariant_param_prints_twice
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

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

System.Action<object> once=v=>{__P((v).ToString()); __P((v).ToString());}; System.Action<string> twice=once; twice("x");
__Check("x\nx");
