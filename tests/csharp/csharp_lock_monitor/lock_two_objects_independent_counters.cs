// vybe-test: csharp/csharp_lock_monitor/lock_two_objects_independent_counters
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

object a = new object();
object b = new object();
int ca = 0;
int cb = 0;
lock (a) { ca++; }
lock (b) { cb += 2; }
__P((ca + cb).ToString());
__Check("3");
