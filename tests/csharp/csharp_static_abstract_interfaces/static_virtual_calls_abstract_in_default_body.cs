// vybe-test: csharp/csharp_static_abstract_interfaces/static_virtual_calls_abstract_in_default_body
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IWrap<T> where T:IWrap<T>{static abstract T Core(); static virtual T Outer(){return Core();}}
struct Core:IWrap<Core>{public int N; public static Core Core()=>new Core{N=4};}
__Check((Core.Outer().N).ToString(), "4");
