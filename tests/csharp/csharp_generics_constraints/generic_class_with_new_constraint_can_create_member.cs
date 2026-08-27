// vybe-test: csharp/csharp_generics_constraints/generic_class_with_new_constraint_can_create_member
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

__P((new Factory<Item>().Build().Name).ToString());
__Check("built");

class Factory<T> where T : new() { public T Build() { return new T(); } }

class Item { public string Name = "built"; }

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
