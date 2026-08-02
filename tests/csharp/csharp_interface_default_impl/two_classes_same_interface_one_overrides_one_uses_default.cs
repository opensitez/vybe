// vybe-test: csharp/csharp_interface_default_impl/two_classes_same_interface_one_overrides_one_uses_default
// origin: languages/csharp/tests/csharp/test_csharp_interface_default_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFormat{string Format(int n)=>$"[{n}]";}
class A:IFormat{}
class B:IFormat{public string Format(int n)=>n.ToString();}
IFormat a=new A(); IFormat b=new B();
__Check((a.Format(5)).ToString(), "[5]");
__Check((b.Format(5)).ToString(), "5");
