// vybe-test: csharp/csharp_class_indexers/indexer_on_readonly_wrapper_exposes_underlying_element
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

using static __Harness;

__P((new ReadWrapper()[0]).ToString());
__Check("5");

class ReadWrapper {
    readonly int[] data = { 5, 6 };
    public int this[int i] { get { return data[i]; } }
}

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
