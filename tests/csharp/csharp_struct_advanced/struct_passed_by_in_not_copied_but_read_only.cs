// vybe-test: csharp/csharp_struct_advanced/struct_passed_by_in_not_copied_but_read_only
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vec{public int X,Y;}
int Sum(in Vec v)=>v.X+v.Y;
var v=new Vec{X=3,Y=4};
__Check((Sum(in v)).ToString(), "7");
