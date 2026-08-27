// vybe-test: csharp/csharp_functional_patterns/unfold_pattern_generates_fibonacci_via_iteration
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Fibs(){
    int a=0,b=1;
    while(true){yield return a; (a,b)=(b,a+b);}
}
var first8=Fibs().Take(8).ToArray();
__P((first8[7]).ToString());
__Check("13");

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
