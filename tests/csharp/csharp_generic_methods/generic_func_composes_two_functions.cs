// vybe-test: csharp/csharp_generic_methods/generic_func_composes_two_functions
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<A,C> Compose<A,B,C>(System.Func<A,B> f,System.Func<B,C> g)=>x=>g(f(x));
var fn=Compose((int x)=>x*2,(int y)=>y+1);
__Check((fn(5)).ToString(), "11");
