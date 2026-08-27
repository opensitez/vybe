// vybe-test: csharp/csharp_access_modifiers/private_field_only_accessible_within_declaring_class
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

using static __Harness;

__P((new Safe().Get()).ToString());
__Check("42");

class Safe{private int secret=42; public int Get()=>secret;}

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
