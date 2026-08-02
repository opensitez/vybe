// vybe-test: csharp/csharp_primary_constructors/primary_constructor_generic_struct
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Wrap<T>(T value) { public T Value => value; }
__Check((new Wrap<string>("hi").Value).ToString(), "hi");
