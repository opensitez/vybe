// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_returns_interface_implementor
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IProvider<T> where T:IProvider<T>{static abstract T Provide();}
class Service:IProvider<Service>{public string Name="svc"; public static Service Provide()=>new Service();}
__Check((Service.Provide().Name).ToString(), "svc");
