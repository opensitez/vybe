// vybe-test: csharp/csharp_modern/default_parameters
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

using static __Harness;

string Greet(string name = "World") {
    return "Hello " + name;
}
__P((Greet()).ToString());
__P((Greet("Alice")).ToString());
__Check("Hello World\nHello Alice");

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
