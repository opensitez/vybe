// vybe-test: csharp/csharp_events_advanced/event_with_two_subscribers_updates_shared_counter
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

using System; class Hub { public event Action Ping; public void Fire() { Ping(); } } int total = 0; var hub = new Hub(); hub.Ping += () => total += 2; hub.Ping += () => total += 3; hub.Fire(); __P((total).ToString());
__Check("5");
