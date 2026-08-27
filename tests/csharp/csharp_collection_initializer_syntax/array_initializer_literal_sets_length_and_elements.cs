// vybe-test: csharp/csharp_collection_initializer_syntax/array_initializer_literal_sets_length_and_elements
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;

var data = new[] { 10, 20, 30 }
;
__P((data.Length).ToString());
__P((data[1]).ToString());
__Check("3\n20");

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
