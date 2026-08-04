// vybe-test: csharp/csharp_events_advanced/multicast_delegate_combines_two_named_handlers
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

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

using System; void First() { __P(("A").ToString()); } void Second() { __P(("B").ToString()); } Action action = First; action += Second; action();
__Check("A\nB");
