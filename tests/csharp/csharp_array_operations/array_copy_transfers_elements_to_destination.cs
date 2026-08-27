// vybe-test: csharp/csharp_array_operations/array_copy_transfers_elements_to_destination
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

using static __Harness;

int[] src = {10,20,30}
;
int[] dst = new int[3];
System.Array.Copy(src, dst, 3);
__P((dst[1]).ToString());
__Check("20");

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
