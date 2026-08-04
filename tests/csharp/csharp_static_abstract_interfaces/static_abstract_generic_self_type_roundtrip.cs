// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_generic_self_type_roundtrip
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

interface ISelf<T> where T:ISelf<T>{static abstract T Identity(T v);}
struct Wrap:ISelf<Wrap>{public int N; public static Wrap Identity(Wrap v)=>v;}
var w=new Wrap{N=3}; __P((Wrap.Identity(w).N).ToString());
__Check("3");
