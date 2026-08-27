// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_three_string_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object p=new Person{A=new Addr{S=new Street{Name="Main"}}}
;
__P((p is Person{A:{S:{Name:"Main"}}}).ToString());
__Check("True");

class Street { public string Name; }

class Addr { public Street S; }

class Person { public Addr A; }

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
