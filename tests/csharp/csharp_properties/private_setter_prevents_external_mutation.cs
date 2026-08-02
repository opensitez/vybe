// vybe-test: csharp/csharp_properties/private_setter_prevents_external_mutation
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    public int Count { get; private set; }
    public void Increment() => Count++;
}
var c = new Counter(); c.Increment(); c.Increment();
__Check((c.Count).ToString(), "2");
