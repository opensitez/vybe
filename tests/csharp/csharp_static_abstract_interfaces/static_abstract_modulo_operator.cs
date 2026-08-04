// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_modulo_operator
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

interface IMod<T> where T:IMod<T>{static abstract T operator%(T a,T b);}
struct Mod:IMod<Mod>{public int V; public static Mod operator%(Mod a,Mod b)=>new Mod{V=a.V%b.V};}
__P(((new Mod{V=10}%new Mod{V=3}).V).ToString());
__Check("1");
