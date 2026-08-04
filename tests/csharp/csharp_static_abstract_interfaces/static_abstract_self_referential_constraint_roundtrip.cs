// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_self_referential_constraint_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

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

interface IRound<T> where T:IRound<T>{static abstract T RoundTrip(T input);}
struct Echo:IRound<Echo>{public int N; public static Echo RoundTrip(Echo input)=>input;}
var e=new Echo{N=12}; __P((Echo.RoundTrip(e).N).ToString());
__Check("12");
