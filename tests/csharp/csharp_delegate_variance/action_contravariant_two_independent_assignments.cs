// vybe-test: csharp/csharp_delegate_variance/action_contravariant_two_independent_assignments
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

System.Action<object> a=v=>__P(("a").ToString()); System.Action<object> b=v=>__P(("b").ToString()); System.Action<string> sa=a; System.Action<string> sb=b; sa(""); sb("");
__Check("a\nb");
