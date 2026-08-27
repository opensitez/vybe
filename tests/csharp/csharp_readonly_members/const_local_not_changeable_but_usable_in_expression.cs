// vybe-test: csharp/csharp_readonly_members/const_local_not_changeable_but_usable_in_expression
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

using static __Harness;

const int MAX=100;
__P((MAX*2).ToString());
__Check("200");

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
