// vybe-test: csharp/csharp_functional_patterns/memoize_caches_expensive_computation
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

using static __Harness;

var cache=new System.Collections.Generic.Dictionary<int,int>();
System.Func<int,int> fib=null;
fib=n=>{
    if(n<=1) return n;
    if(cache.TryGetValue(n,out int v)) return v;
    var r=fib(n-1)+fib(n-2);
    cache[n]=r;
    return r;
}
;
__P((fib(10)).ToString());
__Check("55");

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
