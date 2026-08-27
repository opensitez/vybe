// vybe-test: csharp/collections_advanced/dict_add_and_access
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var dict = new Dictionary<string, int>();
dict.Add("one", 1);
dict.Add("two", 2);
dict.Add("three", 3);
__P((dict["two"]).ToString());
__P((dict.Count).ToString());
__Check("2\n3");

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
