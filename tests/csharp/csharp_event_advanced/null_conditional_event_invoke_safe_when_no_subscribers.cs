// vybe-test: csharp/csharp_event_advanced/null_conditional_event_invoke_safe_when_no_subscribers
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

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

class Btn{public event System.Action Click;}
var b=new Btn();
b.Click?.Invoke();
__P(("ok").ToString());
__Check("ok");
