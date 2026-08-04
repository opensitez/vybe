// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_property_chain_default
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

interface IChain<T> where T:IChain<T>{static virtual string Name=>"base"; static abstract T Instance();}
struct Link:IChain<Link>{public static Link Instance()=>new Link();}
__P((Link.Name).ToString());
__Check("base");
