// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_one_class_overrides_other_uses_default
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IScale{int Scale(int n)=>n;} class Plain:IScale{} class Double:IScale{public int Scale(int n)=>n*2;} __Check((new Plain().Scale(5)+new Double().Scale(5)).ToString(), "15");
