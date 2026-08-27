// vybe-test: csharp/csharp_recursive_algorithms/recursive_gcd_computes_greatest_common_divisor
// origin: languages/csharp/tests/csharp/test_csharp_recursive_algorithms.rs

using static __Harness;

int Gcd(int a,int b)=>b==0?a:Gcd(b,a%b);
__P((Gcd(48,18)).ToString());
__Check("6");

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
