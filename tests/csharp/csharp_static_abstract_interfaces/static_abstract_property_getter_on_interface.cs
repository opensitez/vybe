// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_property_getter_on_interface
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

interface IUnit<T> where T:IUnit<T>{static abstract T Zero{get;}}
struct Counter:IUnit<Counter>{public int V; public static Counter Zero=>new Counter{V=0};}
__P((Counter.Zero.V).ToString());
__Check("0");
