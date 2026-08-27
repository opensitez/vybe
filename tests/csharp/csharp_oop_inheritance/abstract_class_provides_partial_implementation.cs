// vybe-test: csharp/csharp_oop_inheritance/abstract_class_provides_partial_implementation
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

__P((new Impl().Double()).ToString());
__Check("10");

abstract class Base {
    public abstract int Value();
    public int Double() => Value() * 2;
}

class Impl : Base { public override int Value() => 5; }

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
