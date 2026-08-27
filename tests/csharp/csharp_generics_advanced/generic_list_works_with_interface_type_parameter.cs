// vybe-test: csharp/csharp_generics_advanced/generic_list_works_with_interface_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

using static __Harness;

var animals = new System.Collections.Generic.List<IAnimal> { new Cat() }
;
foreach(var a in animals) __P((a.Sound()).ToString());
__Check("meow");

interface IAnimal { string Sound(); }

class Cat : IAnimal { public string Sound() => "meow"; }

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
