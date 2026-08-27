// vybe-test: csharp/collections/list_reverse
// origin: languages/csharp/tests/csharp/test_collections.rs

using static __Harness;

var list = new List<int>();
list.Add(1);
list.Add(2);
list.Add(3);
list.Reverse();
foreach (var x in list) { __P((x).ToString()); }
__Check("3\n2\n1");

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
