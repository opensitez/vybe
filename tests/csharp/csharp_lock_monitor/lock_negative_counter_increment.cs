// vybe-test: csharp/csharp_lock_monitor/lock_negative_counter_increment
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
int counter = -2;
lock (gate) { counter++; }
__P((counter).ToString());
__Check("-1");
