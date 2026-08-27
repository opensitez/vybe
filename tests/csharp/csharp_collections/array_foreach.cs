// vybe-test: csharp/csharp_collections/array_foreach
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

using static __Harness;

string[] names = {"Alice", "Bob", "Carol"}
;
foreach (var name in names) {
    __P((name).ToString());
}
__Check("Alice\nBob\nCarol");

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
