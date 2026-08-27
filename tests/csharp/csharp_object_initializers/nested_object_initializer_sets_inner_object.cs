// vybe-test: csharp/csharp_object_initializers/nested_object_initializer_sets_inner_object
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

using static __Harness;

var p=new Person{Name="Bob",Home=new Address{City="Paris"}}
;
__P((p.Home.City).ToString());
__Check("Paris");

class Address{public string City;}

class Person{public string Name;public Address Home;}

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
