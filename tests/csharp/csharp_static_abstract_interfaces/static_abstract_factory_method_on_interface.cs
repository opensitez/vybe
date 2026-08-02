// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_factory_method_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFactory<T> where T:IFactory<T>{static abstract T Create(int n);}
struct Widget:IFactory<Widget>{public int V; public static Widget Create(int n)=>new Widget{V=n};}
__Check((Widget.Create(5).V).ToString(), "5");
