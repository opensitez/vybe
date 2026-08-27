// vybe-test: csharp/csharp_abstract_class/abstract_property_overridden_in_concrete_class
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

using static __Harness;

Derived d = new Derived();
d.Title = "Sample";
__P(d.Title);
__Check("Sample");

abstract class Base {
    public abstract string Title { get; set; }
}
class Derived : Base {
    public override string Title { get; set; } = "";
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
