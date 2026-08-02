// vybe-test: csharp/csharp_constructor_chains/generic_class_constructor_can_store_type_specific_value
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T> { T value; public Box(T value) { this.value = value; } public T Read() { return value; } } __Check((new Box<string>("text").Read()).ToString(), "text");
