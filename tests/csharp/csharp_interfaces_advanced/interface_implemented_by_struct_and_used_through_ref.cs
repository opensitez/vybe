// vybe-test: csharp/csharp_interfaces_advanced/interface_implemented_by_struct_and_used_through_ref
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface IArea{double Area();}
struct Rect:IArea{public double W,H; public double Area()=>W*H;}
IArea a=new Rect{W=3,H=4};
__P((a.Area()).ToString());
__Check("12");
