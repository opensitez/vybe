// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_self_referential_constraint_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRound<T> where T:IRound<T>{static abstract T RoundTrip(T input);}
struct Echo:IRound<Echo>{public int N; public static Echo RoundTrip(Echo input)=>input;}
var e=new Echo{N=12}; __Check((Echo.RoundTrip(e).N).ToString(), "12");
