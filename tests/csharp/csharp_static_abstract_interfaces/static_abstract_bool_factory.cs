// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_bool_factory
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

interface IFlag<T> where T:IFlag<T>{static abstract T True(); static abstract T False();}
struct Bit:IFlag<Bit>{public bool On; public static Bit True()=>new Bit{On=true}; public static Bit False()=>new Bit{On=false};}
__P((Bit.True().On).ToString());
__Check("True");
