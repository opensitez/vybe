// vybe-test: csharp/csharp_nested_classes/nested_class_can_implement_interface
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IValue{int Get();}
class Host{public class Impl:IValue{public int Get()=>5;}}
IValue v=new Host.Impl();
__Check((v.Get()).ToString(), "5");
