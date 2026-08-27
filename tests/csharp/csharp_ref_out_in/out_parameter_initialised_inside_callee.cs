// vybe-test: csharp/csharp_ref_out_in/out_parameter_initialised_inside_callee
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

using static __Harness;

void Minmax(int[] a, out int min, out int max){
    min=a[0]; max=a[0];
    foreach(var v in a){if(v<min)min=v; if(v>max)max=v;}
}
Minmax(new[]{3,1,4,1,5,9}, out int lo, out int hi);
__P((lo).ToString());
__P((hi).ToString());
__Check("1\n9");

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
