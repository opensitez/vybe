// vybe-test: csharp/csharp_classes/class_this_reference
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

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
    private int count = 0;
    public void Increment() { this.count++; }
    public int GetCount() { return this.count; }
}
var c = new Counter();
c.Increment();
c.Increment();
c.Increment();
__P((c.GetCount()).ToString());
__Check("3");
