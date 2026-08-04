// vybe-test: csharp/csharp_initialization_order/static_constructor_runs_once_before_first_instance_creation
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

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
    static Counter() { __P(("static-ctor").ToString()); }
    public Counter() { __P(("instance").ToString()); }
}
new Counter();
new Counter();
__Check("static-ctor\ninstance\ninstance");
