// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_on_interface_with_multiple_members
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IOps{int Add(int a,int b)=>a+b; int Mul(int a,int b)=>a*b;} class Ops:IOps{} var o=new Ops(); __Check((o.Add(2,3)+o.Mul(2,3)).ToString(), "11");
