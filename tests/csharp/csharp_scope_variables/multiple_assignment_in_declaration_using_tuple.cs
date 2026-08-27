// vybe-test: csharp/csharp_scope_variables/multiple_assignment_in_declaration_using_tuple
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

using static __Harness;

var (a, b) = (3, 7);
__P((a).ToString());
__P((b).ToString());
__Check("3\n7");

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
