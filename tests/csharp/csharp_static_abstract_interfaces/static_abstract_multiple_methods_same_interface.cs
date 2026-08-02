// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_multiple_methods_same_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBoth<T> where T:IBoth<T>{static abstract T FromInt(int n); static abstract T FromString(string s);}
struct Dual:IBoth<Dual>{public string Text; public static Dual FromInt(int n)=>new Dual{Text=n.ToString()}; public static Dual FromString(string s)=>new Dual{Text=s};}
__Check((Dual.FromString("ok").Text).ToString(), "ok");
