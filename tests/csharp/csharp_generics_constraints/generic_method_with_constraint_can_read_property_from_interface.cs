// vybe-test: csharp/csharp_generics_constraints/generic_method_with_constraint_can_read_property_from_interface
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

string Read<T>(T item) where T : INamed { return item.Name; }
__P((Read(new User())).ToString());
__Check("Grace");

interface INamed { string Name { get; } }

class User : INamed { public string Name => "Grace"; }

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
