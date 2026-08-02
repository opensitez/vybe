// vybe-test: csharp/csharp_modern/default_parameters
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Greet(string name = "World") {
    return "Hello " + name;
}
__Check((Greet()).ToString(), "Hello World");
__Check((Greet("Alice")).ToString(), "Hello Alice");
