// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_method_default_calls_primary
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

interface IDouble<T> where T:IDouble<T>{static abstract T One(); static virtual T Two(){return One();}}
struct Dup:IDouble<Dup>{public int V; public static Dup One()=>new Dup{V=1}; public static Dup Two()=>new Dup{V=2};}
__P((Dup.Two().V).ToString());
__Check("2");
