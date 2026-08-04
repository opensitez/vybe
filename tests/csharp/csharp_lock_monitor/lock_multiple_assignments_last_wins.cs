// vybe-test: csharp/csharp_lock_monitor/lock_multiple_assignments_last_wins
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

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

object gate = new object();
int counter = 0;
lock (gate) {
    counter = 1;
    counter = 2;
    counter = 9;
}
__P((counter).ToString());
__Check("9");
