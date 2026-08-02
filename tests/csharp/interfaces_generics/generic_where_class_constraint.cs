// vybe-test: csharp/interfaces_generics/generic_where_class_constraint
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Container<T> where T : class {
    public T Value;
    public bool IsNull() { return Value == null; }
}
var c = new Container<string>();
__Check((c.IsNull()).ToString(), "True");
c.Value = "hello";
__Check((c.IsNull()).ToString(), "False");
