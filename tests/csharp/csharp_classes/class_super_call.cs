// vybe-test: csharp/csharp_classes/class_super_call
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

var d = new Derived();
__P((d.Greet()).ToString());
__Check("Hello World");

class Base {
    public virtual string Greet() { return "Hello"; }
}

class Derived : Base {
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
