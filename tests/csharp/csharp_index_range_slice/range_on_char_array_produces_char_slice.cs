// vybe-test: csharp/csharp_index_range_slice/range_on_char_array_produces_char_slice
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

using static __Harness;

char[] chars = new char[] { 'a', 'b', 'c' };
char[] slice = chars[0..2];
__P(slice.Length.ToString());
__Check("2");
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
