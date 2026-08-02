// vybe-test: csharp/csharp_default_interface_methods_deep/default_interface_property_setter_pair
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ICounter{int Count{get;set;} void Inc(){Count++;}} class C:ICounter{public int Count{get;set;}} var c=new C(); c.Inc(); __Check((c.Count).ToString(), "1");
