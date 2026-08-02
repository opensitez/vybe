// vybe-test: csharp/csharp_default_interface_methods_deep/default_interface_method_class_override_via_interface_typed
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFormat{string Show(int n)=>n.ToString();} class Custom:IFormat{public string Show(int n)=>"x"+n;} IFormat f=new Custom(); __Check((f.Show(3)).ToString(), "x3");
