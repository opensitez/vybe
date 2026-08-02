// vybe-test: csharp/csharp_functional_patterns/memoize_caches_expensive_computation
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var cache=new System.Collections.Generic.Dictionary<int,int>();
System.Func<int,int> fib=null;
fib=n=>{
    if(n<=1) return n;
    if(cache.TryGetValue(n,out int v)) return v;
    var r=fib(n-1)+fib(n-2);
    cache[n]=r;
    return r;
};
__Check((fib(10)).ToString(), "55");
