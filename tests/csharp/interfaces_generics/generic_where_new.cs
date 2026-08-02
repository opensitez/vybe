// vybe-test: csharp/interfaces_generics/generic_where_new
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Factory<T> where T : new() {
    public T Create() { return new T(); }
}
class Item {
    public string Name = "default";
}
var f = new Factory<Item>();
var item = f.Create();
__Check((item.Name).ToString(), "default");
