// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_flags_in_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P(((int)new Auth().All()).ToString());
__Check("3");

class Auth{[System.Flags] public enum Perm{None=0,Read=1,Write=2} public Perm All()=>Perm.Read|Perm.Write;}

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
