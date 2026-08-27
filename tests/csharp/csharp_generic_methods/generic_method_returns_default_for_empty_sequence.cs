// vybe-test: csharp/csharp_generic_methods/generic_method_returns_default_for_empty_sequence
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

using static __Harness;

T FirstOrDefault<T>(T[] arr)=>arr.Length>0?arr[0]:default;
__P((FirstOrDefault(new int[]{})).ToString());
__P((FirstOrDefault(new[]{9})).ToString());
__Check("0\n9");

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
