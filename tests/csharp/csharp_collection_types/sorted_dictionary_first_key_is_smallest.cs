// vybe-test: csharp/csharp_collection_types/sorted_dictionary_first_key_is_smallest
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

using static __Harness;

var sd=new System.Collections.Generic.SortedDictionary<int,string>{{3,"c"},{1,"a"},{2,"b"}}
;
__P((sd.Keys.First()).ToString());
__Check("1");

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
