// vybe-test: csharp/csharp_array_advanced/array_convert_all_transforms_each_element
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

using static __Harness;

int[] src={1,2,3}
;
string[] dst=System.Array.ConvertAll(src,n=>n.ToString()+"x");
__P((dst[1]).ToString());
__Check("2x");

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
