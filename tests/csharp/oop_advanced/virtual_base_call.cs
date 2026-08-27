// vybe-test: csharp/oop_advanced/virtual_base_call
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var c = new Child();
__P((c.Greet()).ToString());
__Check("Hello World");

class Base {
    public virtual string Greet() { return "Hello"; }
}

class Child : Base {
    public override string Greet() { return base.Greet() + " World"; }
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
