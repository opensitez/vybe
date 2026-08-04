// vybe-test: csharp/classes/class_multiple_instances
// origin: languages/csharp/tests/csharp/test_classes.rs

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
            int count;
            public Counter(int start) { this.count = start; }
            public void Inc() { this.count = this.count + 1; }
            public int Get() { return this.count; }
        }
        var a = new Counter(0);
        var b = new Counter(100);
        a.Inc(); a.Inc();
        b.Inc();
        __P((a.Get()).ToString());
        __P((b.Get()).ToString());
__Check("2\n101");
