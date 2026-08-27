// vybe-test: csharp/csharp_anonymous_types/anonymous_type_created_with_new_projection_syntax
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

using static __Harness;

var a=new{Name="Alice",Age=30}
;
__P((a.Name).ToString());
__P((a.Age).ToString());
__Check("Alice\n30");

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
