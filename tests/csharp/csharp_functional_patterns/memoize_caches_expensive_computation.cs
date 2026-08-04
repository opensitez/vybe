// vybe-test: csharp/csharp_functional_patterns/memoize_caches_expensive_computation
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
__P((fib(10)).ToString());
__Check("55");
