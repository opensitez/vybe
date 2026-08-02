// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_interface_hiding_with_new_class_method
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IHide{int V()=>1;} class Hide:IHide{public new int V()=>2;} __Check((new Hide().V()).ToString(), "2");
