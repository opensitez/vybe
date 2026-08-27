// vybe-test: csharp/csharp_immutable_collections/immutable_array_indexer_reads_element
// origin: languages/csharp/tests/csharp/test_csharp_immutable_collections.rs

using static __Harness;

var arr=System.Collections.Immutable.ImmutableArray.Create(10,20,30);
__P((arr[1]).ToString());
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
