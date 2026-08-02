// vybe-test: csharp/csharp_properties_accessors/private_setter_property_changes_through_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((counter.Value).ToString(), "2");
