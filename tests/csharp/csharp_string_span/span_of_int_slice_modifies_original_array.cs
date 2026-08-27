// vybe-test: csharp/csharp_string_span/span_of_int_slice_modifies_original_array
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

using static __Harness;

int[] arr={1,2,3,4,5}
;
System.Span<int> s=arr.AsSpan(1,3);
s[0]=99;
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
