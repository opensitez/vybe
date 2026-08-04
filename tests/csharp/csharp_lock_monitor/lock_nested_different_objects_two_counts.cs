// vybe-test: csharp/csharp_lock_monitor/lock_nested_different_objects_two_counts
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

object outer = new object();
object inner = new object();
int counter = 0;
lock (outer) {
    counter++;
    lock (inner) { counter++; }
}
__P((counter).ToString());
__Check("2");
