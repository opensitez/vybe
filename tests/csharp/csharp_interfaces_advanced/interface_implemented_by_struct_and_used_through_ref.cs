// vybe-test: csharp/csharp_interfaces_advanced/interface_implemented_by_struct_and_used_through_ref
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IArea{double Area();}
struct Rect:IArea{public double W,H; public double Area()=>W*H;}
IArea a=new Rect{W=3,H=4};
__Check((a.Area()).ToString(), "12");
