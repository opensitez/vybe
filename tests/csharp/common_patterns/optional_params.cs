// vybe-test: csharp/common_patterns/optional_params
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Greeter {
    public static string Hello(string name, string greeting = "Hello") {
        return greeting + ", " + name + "!";
    }
}
__P((Greeter.Hello("Alice")).ToString());
__P((Greeter.Hello("Bob", "Hi")).ToString());
__Check("Hello, Alice!\nHi, Bob!");
