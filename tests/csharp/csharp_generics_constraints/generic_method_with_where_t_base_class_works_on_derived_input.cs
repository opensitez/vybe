// vybe-test: csharp/csharp_generics_constraints/generic_method_with_where_t_base_class_works_on_derived_input
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

string Read<T>(T person) where T : Person { return person.Name; }
__P((Read(new Admin())).ToString());
__Check("Ada");

class Person { public string Name = "Ada"; }

class Admin : Person { }

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
