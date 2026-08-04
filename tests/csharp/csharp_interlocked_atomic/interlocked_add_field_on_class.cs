// vybe-test: csharp/csharp_interlocked_atomic/interlocked_add_field_on_class
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic.rs

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

class Counter {
    public int Value = 5;
    public void Add(int n) { System.Threading.Interlocked.Add(ref Value, n); }
}
var c = new Counter();
c.Add(3);
__P((c.Value).ToString());
__Check("8");
