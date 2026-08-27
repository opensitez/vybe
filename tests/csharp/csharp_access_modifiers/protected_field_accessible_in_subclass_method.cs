// vybe-test: csharp/csharp_access_modifiers/protected_field_accessible_in_subclass_method
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

using static __Harness;

__P((new B().Read()).ToString());
__Check("7");

class A{protected int Value=7;}

class B:A{public int Read()=>Value;}

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
