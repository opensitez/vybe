// vybe-test: csharp/csharp_ref_out_in/ref_local_aliases_array_element
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

using static __Harness;

int[] arr={1,2,3}
;
ref int second=ref arr[1];
second=99;
__P((arr[1]).ToString());
__Check("99");

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
