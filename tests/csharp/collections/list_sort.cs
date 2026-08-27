// vybe-test: csharp/collections/list_sort
// origin: languages/csharp/tests/csharp/test_collections.rs

using static __Harness;

var list = new List<int>();
list.Add(3);
list.Add(1);
list.Add(2);
list.Sort();
foreach (var x in list) { __P((x).ToString()); }
__Check("1\n2\n3");

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
