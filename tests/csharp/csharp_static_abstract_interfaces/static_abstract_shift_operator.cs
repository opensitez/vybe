// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_shift_operator
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

interface IShift<T> where T:IShift<T>{static abstract T operator<<(T v,int n);}
struct Bits:IShift<Bits>{public int V; public static Bits operator<<(Bits v,int n)=>new Bits{V=v.V<<n};}
__P(((new Bits{V=1}<<3).V).ToString());
__Check("8");
