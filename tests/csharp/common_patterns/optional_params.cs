// vybe-test: csharp/common_patterns/optional_params
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P((Greeter.Hello("Alice")).ToString());
__P((Greeter.Hello("Bob", "Hi")).ToString());
__Check("Hello, Alice!\nHi, Bob!");

class Greeter {
    public static string Hello(string name, string greeting = "Hello") {
        return greeting + ", " + name + "!";
    }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
