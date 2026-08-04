// vybe-test: csharp/csharp_delegate_variance/action_contravariant_with_local_function
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

void Show(object o)=>__P((o).ToString()); System.Action<object> baseAct=Show; System.Action<string> derivedAct=baseAct; derivedAct("fn");
__Check("fn");
