// vybe-test: csharp/classes/property_chain
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

class Inner { public int value; public Inner(int v) { this.value = v; } }
        class Outer { public Inner inner; public Outer(int v) { this.inner = new Inner(v); } }
        var o = new Outer(42);
        __P((o.inner.value).ToString());
__Check("42");
