// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_factory_method_on_interface
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

interface IFactory<T> where T:IFactory<T>{static abstract T Create(int n);}
struct Widget:IFactory<Widget>{public int V; public static Widget Create(int n)=>new Widget{V=n};}
__P((Widget.Create(5).V).ToString());
__Check("5");
