// vybe-test: csharp/csharp_functional_patterns/partial_application_creates_specialized_function
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

using static __Harness;

System.Func<int,System.Func<int,int>> add=a=>b=>a+b;
var add10=add(10);
__P((add10(5)).ToString());
__P((add10(20)).ToString());
__Check("15\n30");

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
