// vybe-test: csharp/csharp_lock_monitor/lock_class_instance_field_gate
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

class SafeCounter {
    private object gate = new object();
    public int Value = 0;
    public void Add(int n) { lock (gate) { Value += n; } }
}
var sc = new SafeCounter();
sc.Add(3);
sc.Add(4);
__P((sc.Value).ToString());
__Check("7");
