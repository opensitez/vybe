// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_interface_nameof_type
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

interface IName<T> where T:IName<T>{static abstract string TypeName();}
struct Named:IName<Named>{public static string TypeName()=>nameof(Named);}
__P((Named.TypeName()).ToString());
__Check("Named");
