// vybe-test: csharp/csharp_generic_methods/generic_method_swap_exchanges_two_values_via_ref
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

using static __Harness;

void Swap<T>(ref T a,ref T b){T tmp=a;a=b;b=tmp;}
int x=1,y=2;
Swap(ref x,ref y);
__P((x).ToString());
__P((y).ToString());
__Check("2\n1");

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
