// vybe-test: csharp/csharp_new_features/using_static_imports_type_members_without_qualifier
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

using static __Harness;
using static System.Math;

__P((Sqrt(16)).ToString());
__Check("4");

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
