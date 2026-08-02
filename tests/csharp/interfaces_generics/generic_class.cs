// vybe-test: csharp/interfaces_generics/generic_class
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T> {
    public T Value;
    public Box(T val) { Value = val; }
}
var intBox = new Box<int>(42);
var strBox = new Box<string>("hello");
__Check((intBox.Value).ToString(), "42");
__Check((strBox.Value).ToString(), "hello");
