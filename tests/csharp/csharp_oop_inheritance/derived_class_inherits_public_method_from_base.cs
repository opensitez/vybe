// vybe-test: csharp/csharp_oop_inheritance/derived_class_inherits_public_method_from_base
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

__P((new Derived().Hello()).ToString());
__Check("hello");

class Base { public string Hello() => "hello"; }

class Derived : Base { }

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
