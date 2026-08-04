// vybe-test: csharp/csharp_events_advanced/action_delegate_array_invokes_each_entry
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

using System; Action[] actions = { () => __P(("one").ToString()), () => __P(("two").ToString()) }; foreach (var action in actions) action();
__Check("one\ntwo");
