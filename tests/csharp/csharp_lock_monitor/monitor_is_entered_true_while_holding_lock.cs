// vybe-test: csharp/csharp_lock_monitor/monitor_is_entered_true_while_holding_lock
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
int count = 0;
System.Threading.Monitor.Enter(gate);
count = System.Threading.Monitor.IsEntered(gate) ? 1 : 0;
System.Threading.Monitor.Exit(gate);
__P((count).ToString());
__Check("1");
