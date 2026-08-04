// vybe-test: csharp/csharp_fluent_builder_pattern/static_factory_method_can_start_chain_on_new_instance
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

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
    int value;
    public static Counter Start(int seed) {
        var counter = new Counter();
        counter.value = seed;
        return counter;
    }
    public Counter Bump() { value++; return this; }
    public int Read() { return value; }
}
__P((Counter.Start(10).Bump().Bump().Read()).ToString());
__Check("12");
