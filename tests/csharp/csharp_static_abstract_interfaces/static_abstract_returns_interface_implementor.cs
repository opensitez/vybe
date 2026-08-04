// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_returns_interface_implementor
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

interface IProvider<T> where T:IProvider<T>{static abstract T Provide();}
class Service:IProvider<Service>{public string Name="svc"; public static Service Provide()=>new Service();}
__P((Service.Provide().Name).ToString());
__Check("svc");
