// vybe-test: csharp/csharp_primary_constructors/primary_constructor_generic_class_with_type_param
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T>(T item) { public T Item => item; }
__Check((new Box<int>(42).Item).ToString(), "42");
