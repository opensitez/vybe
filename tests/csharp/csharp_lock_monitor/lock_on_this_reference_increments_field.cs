// vybe-test: csharp/csharp_lock_monitor/lock_on_this_reference_increments_field
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

class Box {
    public int counter = 0;
    public void Inc() { lock (this) { counter++; } }
}
var box = new Box();
box.Inc();
__P((box.counter).ToString());
__Check("1");
