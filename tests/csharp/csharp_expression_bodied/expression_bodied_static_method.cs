// vybe-test: csharp/csharp_expression_bodied/expression_bodied_static_method
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class Utils{public static int Clamp(int v,int lo,int hi)=>v<lo?lo:v>hi?hi:v;}
__Check((Utils.Clamp(15,0,10)).ToString(), "10");
