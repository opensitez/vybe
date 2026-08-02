// vybe-test: csharp/csharp_primary_constructors/primary_constructor_string_param_used_in_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Greeter(string prefix) {
    public string Greet(string name) => prefix + " " + name;
}
__Check((new Greeter("Hello").Greet("World")).ToString(), "Hello World");
