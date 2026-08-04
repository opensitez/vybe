// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_operator_plus_on_numeric_interface
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

interface IAddable<T> where T:IAddable<T>{static abstract T operator+(T a,T b);}
struct Vec:IAddable<Vec>{public int X; public static Vec operator+(Vec a,Vec b)=>new Vec{X=a.X+b.X};}
__P(((new Vec{X=2}+new Vec{X=3}).X).ToString());
__Check("5");
