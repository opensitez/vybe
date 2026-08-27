// vybe-test: csharp/csharp_string_span/span_copy_to_writes_into_destination
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

using static __Harness;

int[] src={1,2,3}
;
int[] dst=new int[3];
src.AsSpan().CopyTo(dst);
__P((dst[2]).ToString());
__Check("3");

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
