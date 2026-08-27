// vybe-test: csharp/csharp_functional_patterns/method_chaining_builds_fluent_pipeline
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

using static __Harness;

var result=new[]{5,3,8,1,4}
.Where(x=>x>2)
    .OrderBy(x=>x)
    .Select(x=>x*10)
    .First();
__P((result).ToString());
__Check("30");

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
