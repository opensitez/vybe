// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_delegate_declared_in_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new MathUtil.Calc().Run((x,y)=>x+y,2,3)).ToString());
__Check("5");

class MathUtil{public delegate int Op(int a,int b); public class Calc{public int Run(Op f,int a,int b)=>f(a,b);}}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
