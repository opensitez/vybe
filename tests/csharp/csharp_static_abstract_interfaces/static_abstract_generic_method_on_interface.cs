// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_generic_method_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IConvert<T> where T:IConvert<T>{static abstract T From<U>(U value);}
struct Box:IConvert<Box>{public string Text; public static Box From<U>(U value)=>new Box{Text=value.ToString()};}
__Check((Box.From(12).Text).ToString(), "12");
