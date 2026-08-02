// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_accesses_static_class_field
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IStatic{int Read()=>Holder.N;} static class Holder{public static int N=8;} class R:IStatic{} __Check((new R().Read()).ToString(), "8");
