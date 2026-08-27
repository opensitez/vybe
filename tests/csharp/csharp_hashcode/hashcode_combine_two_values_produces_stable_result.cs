// vybe-test: csharp/csharp_hashcode/hashcode_combine_two_values_produces_stable_result
// origin: languages/csharp/tests/csharp/test_csharp_hashcode.rs

using static __Harness;

int h1=System.HashCode.Combine(1,2);
int h2=System.HashCode.Combine(1,2);
__P((h1==h2).ToString());
__Check("True");

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
