// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_bitwise_and_operator
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

interface IBit<T> where T:IBit<T>{static abstract T operator&(T a,T b);}
struct Mask:IBit<Mask>{public int Bits; public static Mask operator&(Mask a,Mask b)=>new Mask{Bits=a.Bits&b.Bits};}
__P(((new Mask{Bits=7}&new Mask{Bits=3}).Bits).ToString());
__Check("3");
