// vybe-test: csharp/csharp_pattern_property/inheritance_property_pattern_on_derived_type
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Dog{Kind="pet",Legs=4}
;
__P((o is Dog{Legs:4,Kind:"pet"}).ToString());
__Check("True");

class Animal { public string Kind; }

class Dog : Animal { public int Legs; }

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
