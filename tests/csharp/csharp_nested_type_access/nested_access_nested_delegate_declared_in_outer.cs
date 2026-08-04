// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_delegate_declared_in_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class MathUtil{public delegate int Op(int a,int b); public class Calc{public int Run(Op f,int a,int b)=>f(a,b);}} __P((new MathUtil.Calc().Run((x,y)=>x+y,2,3)).ToString());
__Check("5");
