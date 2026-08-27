// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_rejects_wrong_inner
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object p=new Person{Home=new Address{City="Paris"}}
;
__P((p is Person{Home:{City:"London"}}).ToString());
__Check("False");

class Address { public string City; }

class Person { public Address Home; }

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
