// vybe-test: csharp/csharp_lock_monitor/lock_switch_case_increments_matching
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
int code = 2;
lock (gate) {
    switch (code) {
        case 1: counter = 10; break;
        case 2: counter = 20; break;
        default: counter = 0; break;
    }
}
__P((counter).ToString());
__Check("20");
