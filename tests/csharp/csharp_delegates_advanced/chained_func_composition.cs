// vybe-test: csharp/csharp_delegates_advanced/chained_func_composition
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

using static __Harness;

System.Func<int,int> double_it=x=>x*2;
System.Func<int,int> add_three=x=>x+3;
System.Func<int,int> combined=x=>add_three(double_it(x));
__P((combined(5)).ToString());
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
