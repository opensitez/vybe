// vybe-test: csharp/csharp_default_interface_methods_deep/default_interface_method_visible_through_interface_reference
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ICalc{int Double(int n)=>n*2;} class Worker:ICalc{} ICalc w=new Worker(); __Check((w.Double(4)).ToString(), "8");
