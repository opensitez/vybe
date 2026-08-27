// vybe-test: csharp/csharp_generic_constraints/multiple_constraints_combine_with_comma_syntax
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

using static __Harness;

T Make<T>() where T : IName, new() => new T();
__P((Make<Item>().Name()).ToString());
__Check("item");

interface IName { string Name(); }

class Item : IName { public string Name() => "item"; }

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
