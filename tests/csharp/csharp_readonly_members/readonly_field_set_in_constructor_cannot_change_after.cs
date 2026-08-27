// vybe-test: csharp/csharp_readonly_members/readonly_field_set_in_constructor_cannot_change_after
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

using static __Harness;

var obj=new Immutable(42);
__P((obj.Value).ToString());
__Check("42");

class Immutable{public readonly int Value; public Immutable(int v){Value=v;}}

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
