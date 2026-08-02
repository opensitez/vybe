// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_property_chain_default
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IChain<T> where T:IChain<T>{static virtual string Name=>"base"; static abstract T Instance();}
struct Link:IChain<Link>{public static Link Instance()=>new Link();}
__Check((Link.Name).ToString(), "base");
