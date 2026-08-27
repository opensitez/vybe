// vybe-test: csharp/csharp_oop_inheritance/protected_member_is_accessible_in_derived_class
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

__P((new Child().Get()).ToString());
__Check("42");

class Base { protected int Secret = 42; }

class Child : Base { public int Get() => Secret; }

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
