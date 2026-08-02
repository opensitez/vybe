// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_not_visible_as_class_member_without_interface
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IHidden{int Secret()=>9;} class Worker:IHidden{} IHidden w=new Worker(); __Check((w.Secret()).ToString(), "9");
