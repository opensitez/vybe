// vybe-test: csharp/csharp_oop_inheritance/cast_to_base_succeeds_from_derived_instance
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

Base b = (Base)new Derived();
__P((b.X).ToString());
__Check("1");

class Base { public int X = 1; }

class Derived : Base { public int Y = 2; }

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
