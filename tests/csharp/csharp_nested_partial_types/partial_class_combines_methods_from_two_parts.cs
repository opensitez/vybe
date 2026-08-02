// vybe-test: csharp/csharp_nested_partial_types/partial_class_combines_methods_from_two_parts
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Worker {
    public string First() { return "one"; }
}
partial class Worker {
    public string Second() { return "two"; }
}
var worker = new Worker();
__Check((worker.First()).ToString(), "one");
__Check((worker.Second()).ToString(), "two");
