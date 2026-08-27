// vybe-test: csharp/csharp_generics_constraints/generic_method_with_base_class_constraint_accesses_base_member
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

string Read<T>(T value) where T : Base { return value.Name; }
__P((Read(new Child())).ToString());
__Check("base");

class Base { public string Name = "base"; }

class Child : Base { }

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
