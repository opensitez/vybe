// vybe-test: csharp/csharp_exceptions_flow/finally_always_runs_even_after_return
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_flow.rs

using static __Harness;

bool ran=false;
int Compute(){
    try{return 42;}
    finally{ran=true;}
}
int v=Compute();
__P((v).ToString());
__P((ran).ToString());
__Check("42\nTrue");

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
