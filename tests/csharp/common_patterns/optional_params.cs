// vybe-test: csharp/common_patterns/optional_params
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Greeter {
    public static string Hello(string name, string greeting = "Hello") {
        return greeting + ", " + name + "!";
    }
}
__Check((Greeter.Hello("Alice")).ToString(), "Hello, Alice!");
__Check((Greeter.Hello("Bob", "Hi")).ToString(), "Hi, Bob!");
