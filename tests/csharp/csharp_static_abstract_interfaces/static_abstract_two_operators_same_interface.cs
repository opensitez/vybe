// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_two_operators_same_interface
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

interface IOps<T> where T:IOps<T>{static abstract T operator+(T a,T b); static abstract T operator*(T a,int k);}
struct Scale:IOps<Scale>{public int V; public static Scale operator+(Scale a,Scale b)=>new Scale{V=a.V+b.V}; public static Scale operator*(Scale a,int k)=>new Scale{V=a.V*k};}
__P(((new Scale{V=2}*3).V).ToString());
__Check("6");
