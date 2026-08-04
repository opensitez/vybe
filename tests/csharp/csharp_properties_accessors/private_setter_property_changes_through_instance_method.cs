// vybe-test: csharp/csharp_properties_accessors/private_setter_property_changes_through_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

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
    public int Value { get; private set; }
    public void Increment() { Value++; }
}
var counter = new Counter();
counter.Increment();
counter.Increment();
__P((counter.Value).ToString());
__Check("2");
