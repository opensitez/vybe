// vybe-test: csharp/csharp_oop_inheritance/method_hiding_with_new_keyword_selects_by_static_type
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

Parent p = new Child();
__P((p.Name()).ToString());
__Check("Parent");

class Parent { public string Name() => "Parent"; }

class Child : Parent { public new string Name() => "Child"; }

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
