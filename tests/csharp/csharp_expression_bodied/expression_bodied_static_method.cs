// vybe-test: csharp/csharp_expression_bodied/expression_bodied_static_method
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

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

static class Utils{public static int Clamp(int v,int lo,int hi)=>v<lo?lo:v>hi?hi:v;}
__P((Utils.Clamp(15,0,10)).ToString());
__Check("10");
