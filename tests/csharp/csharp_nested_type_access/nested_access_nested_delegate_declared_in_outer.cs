// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_delegate_declared_in_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MathUtil{public delegate int Op(int a,int b); public class Calc{public int Run(Op f,int a,int b)=>f(a,b);}} __Check((new MathUtil.Calc().Run((x,y)=>x+y,2,3)).ToString(), "5");
