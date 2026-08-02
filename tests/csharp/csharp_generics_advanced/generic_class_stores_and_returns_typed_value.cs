// vybe-test: csharp/csharp_generics_advanced/generic_class_stores_and_returns_typed_value
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T> { public T Value; }
var b = new Box<int> { Value = 42 };
__Check((b.Value).ToString(), "42");
