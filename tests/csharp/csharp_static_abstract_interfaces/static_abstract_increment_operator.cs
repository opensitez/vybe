// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_increment_operator
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

interface IInc<T> where T:IInc<T>{static abstract T operator++(T v);}
struct Num:IInc<Num>{public int N; public static Num operator++(Num v)=>new Num{N=v.N+1};}
__P(((++new Num{N=4}).N).ToString());
__Check("5");
