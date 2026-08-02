// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_list_count_helper
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ISize{int Len(System.Collections.Generic.List<int> xs)=>xs.Count;} class Measurer:ISize{} __Check((new Measurer().Len(new System.Collections.Generic.List<int>{1,2,3})).ToString(), "3");
