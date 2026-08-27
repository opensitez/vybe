// vybe-test: csharp/csharp_local_functions/local_function_declared_and_called_within_method
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

using static __Harness;

int Square(int n){
    int Sq(int x)=>x*x;
    return Sq(n);
}
__P((Square(5)).ToString());
__Check("25");

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
