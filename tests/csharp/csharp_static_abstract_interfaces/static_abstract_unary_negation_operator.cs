// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_unary_negation_operator
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

interface INegatable<T> where T:INegatable<T>{static abstract T operator-(T v);}
struct Signed:INegatable<Signed>{public int N; public static Signed operator-(Signed v)=>new Signed{N=-v.N};}
__P(((-new Signed{N=8}).N).ToString());
__Check("-8");
