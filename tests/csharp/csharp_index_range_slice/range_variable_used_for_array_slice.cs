// vybe-test: csharp/csharp_index_range_slice/range_variable_used_for_array_slice
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

using static __Harness;

int[] data={2,4,6,8}
;
System.Range r=new System.Range(1,3);
var slice=data[r];
__P((slice.Length).ToString());
__P((slice[1]).ToString());
__Check("2\n6");

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
