// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_signed_magnitude
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

interface ISign<T> where T:ISign<T>{static abstract T Negate(T v); static abstract T Abs(T v);}
struct IntWrap:ISign<IntWrap>{public int N; public static IntWrap Negate(IntWrap v)=>new IntWrap{N=-v.N}; public static IntWrap Abs(IntWrap v)=>new IntWrap{N=v.N<0?-v.N:v.N};}
__P((IntWrap.Abs(new IntWrap{N=-4}).N).ToString());
__Check("4");
