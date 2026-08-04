// vybe-test: csharp/csharp_lock_monitor/lock_while_loop_increments_to_limit
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
int n = 6;
while (n > 0) {
    lock (gate) { counter++; }
    n--;
}
__P((counter).ToString());
__Check("6");
