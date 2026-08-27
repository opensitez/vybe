// vybe-test: csharp/collections/list_add_numbers
// origin: languages/csharp/tests/csharp/test_collections.rs

using static __Harness;

var list = new List<int>();
list.Add(10);
list.Add(20);
list.Add(30);
var sum = 0;
foreach (var x in list) { sum = sum + x; }
__P((sum).ToString());
__Check("60");

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
