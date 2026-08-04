// vybe-test: csharp/csharp_delegate_types/delegate_null_check_before_invoke_prevents_null_reference
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

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

System.Action handler = null;
handler?.Invoke();
__P(("safe").ToString());
__Check("safe");
